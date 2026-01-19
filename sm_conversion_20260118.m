%% ========================================================================
% sm_conversion_20260112_FIXED.m
% (ALL-IN-ONE, Step0–10)
%
% ✅ All inputs/outputs/process files in ONE folder:
%    D:\Users\jiaweiChiang\Desktop\Supplymentary_code
%
% ✅ NEW CHANGE (your current filenames):
%   - tpl_file fixed in ROOT: "Studies and Fluxes-unit.xlsx"
%     (sheet must contain Paper_number, Study_number, ρb, φ, φ_unit, ρs)
%   - move_template_SM.xlsx is READ-ONLY (DO NOT modify / DO NOT write)
%   - Step10 write-back target: "Studies and Fluxes.xlsx"
%     priority columns: SM_m1..SM_m12 -> m1..m12 -> Jan..Dec (no new columns)
%
% ✅ Only Step10 writes to disk (IN-PLACE overwrite target flux file)
%    (no other .mat / process Excel outputs)
%
% ✅ FIXED:
%   - Step6 DOY detection is robust (not fixed col=5)
%   - Step10 writes ONLY 12 columns back using writecell (NO writetable full overwrite)
% ========================================================================

clear;

%% ===================== ROOT（唯一工作目录） =====================
ROOT = 'D:\Users\jiaweiChiang\Desktop\Supplymentary_code';
if ~exist(ROOT,'dir'); mkdir(ROOT); end
cd(ROOT);

%% ===================== 输入文件（必须在 ROOT 内） =====================
tpl_file    = fullfile(ROOT, 'Studies and Fluxes-unit.xlsx');         % fixed
flux_file   = fullfile(ROOT, 'CH4 uptake data_CH4 FLUX_1_1410.xlsx'); % raw flux
move_file   = fullfile(ROOT, 'move_template_SM.xlsx');               % READ ONLY
target_flux = fullfile(ROOT, 'Studies and Fluxes.xlsx');             % IN-PLACE overwrite

%% ===================== Step 0: 输入检查 =====================
fprintf('\n📌 Step 0: 输入检查...\n');
assert(isfile(tpl_file), '❌ 找不到模板文件：%s', tpl_file);
assert(isfile(flux_file), '❌ 找不到 raw flux 文件：%s', flux_file);
assert(isfile(move_file), '❌ 找不到 move_template_SM.xlsx：%s', move_file);
assert(isfile(target_flux),'❌ 找不到写回目标文件：%s', target_flux);
fprintf('✅ tpl : %s\n', tpl_file);
fprintf('✅ rawflux : %s\n', flux_file);
fprintf('✅ move : %s (READ ONLY)\n', move_file);
fprintf('✅ target : %s\n', target_flux);

%% ===================== Step 1: 从模板提取 & 标准化（仅内存） =====================
fprintf('\n📌 Step 1: 从模板提取物性参数（ρb/φ/ρs + SM_annual_1/2）...\n');

need_cols = ["Paper_number","Study_number","ρb","φ","φ_unit","ρs"];

try
    tpl_sheets = sheetnames(tpl_file);
catch
    [~, tpl_sheets] = xlsfinfo(tpl_file);
end
assert(~isempty(tpl_sheets), '❌ 模板没有可读 sheet：%s', tpl_file);

picked_sheet = '';
tplT = table();
for s = 1:numel(tpl_sheets)
    T = readtable(tpl_file, 'Sheet', tpl_sheets{s}, 'VariableNamingRule','preserve', 'UseExcel', false);
    if all(ismember(need_cols, string(T.Properties.VariableNames)))
        tplT = T;
        picked_sheet = tpl_sheets{s};
        break;
    end
end
assert(~isempty(picked_sheet), '❌ 模板中未找到同时包含列：%s', strjoin(need_cols, ', '));
fprintf('✅ 使用模板 sheet：%s\n', picked_sheet);

toD = @(x) str2double(string(x));
toS = @(x) string(x);

PAPER = toD(tplT.("Paper_number"));
STUDY = toD(tplT.("Study_number"));

% SM_annual_1/2（文本原样，可选）
if any(strcmpi("SM_annual_1", tplT.Properties.VariableNames))
    SMa1_txt = string(tplT.("SM_annual_1"));
else
    SMa1_txt = strings(size(PAPER)); SMa1_txt(:) = missing;
end
if any(strcmpi("SM_annual_2", tplT.Properties.VariableNames))
    SMa2_txt = string(tplT.("SM_annual_2"));
else
    SMa2_txt = strings(size(PAPER)); SMa2_txt(:) = missing;
end

rb    = toD(tplT.("ρb"));      % g/cm^3
phi   = toD(tplT.("φ"));       % (unit per φ_unit)
phi_u = toS(tplT.("φ_unit"));
rs    = toD(tplT.("ρs"));      % g/cm^3

% ρb 合理范围过滤
rb((rb < 0.01) | (rb > 2.65)) = NaN;

% φ -> %（必须用 local_phi_to_pct）
phi_pct = arrayfun(@(v,u) local_phi_to_pct(v,u), phi, phi_u);

% ρs 缺失默认 2.65
rs(isnan(rs)) = 2.65;

tpl_phys = table(PAPER, STUDY, SMa1_txt, SMa2_txt, rb, phi_pct, rs, ...
    'VariableNames', {'Paper_number','Study_number','SM_annual_1','SM_annual_2','rho_b_g_cm3','phi_pct','rho_s_g_cm3'});

% 构建 Map：key="paper|study"（若重复 key，保留首次 stable）
valid_key = isfinite(tpl_phys.Paper_number) & isfinite(tpl_phys.Study_number);
keys_all = cellfun(@(a,b) sprintf('%g|%g', a, b), ...
    num2cell(tpl_phys.Paper_number(valid_key)), num2cell(tpl_phys.Study_number(valid_key)), ...
    'UniformOutput', false);

s1 = string(tpl_phys.SM_annual_1(valid_key)); s1(ismissing(s1)) = "";
s2 = string(tpl_phys.SM_annual_2(valid_key)); s2(ismissing(s2)) = "";
rb_all  = tpl_phys.rho_b_g_cm3(valid_key);
phi_all = tpl_phys.phi_pct(valid_key);
rs_all  = tpl_phys.rho_s_g_cm3(valid_key);

[uniq_key, ia] = unique(keys_all, 'stable');

map_sma1 = containers.Map(uniq_key, cellstr(s1(ia)));
map_sma2 = containers.Map(uniq_key, cellstr(s2(ia)));
map_rb   = containers.Map(uniq_key, rb_all(ia));
map_phi  = containers.Map(uniq_key, phi_all(ia));
map_rs   = containers.Map(uniq_key, rs_all(ia));

%% ===================== Step 2: 在 raw flux sheet 中 sm 列前插入 5 列（仅内存） =====================
fprintf('\n📌 Step 2: 在 raw flux 文件中 sm 列前插入 5 列（只填 sm 行；仅内存）...\n');

try
    fx_sheets = sheetnames(flux_file);
catch
    [~, fx_sheets] = xlsfinfo(flux_file);
end
assert(~isempty(fx_sheets), '❌ raw flux 文件没有可读 sheet：%s', flux_file);

target_sheet = '';
sm_idx = [];

for s = 1:numel(fx_sheets)
    C0 = readcell(flux_file, 'Sheet', fx_sheets{s});
    if isempty(C0) || size(C0,1) < 1, continue; end
    header0 = string(C0(1,:));
    exact = find(strcmpi(strtrim(header0), "sm"), 1, 'first');
    fuzzy = find(contains(lower(header0), "sm"), 1, 'first');
    idx = [];
    if ~isempty(exact), idx = exact;
    elseif ~isempty(fuzzy), idx = fuzzy;
    end
    if ~isempty(idx)
        target_sheet = fx_sheets{s};
        sm_idx = idx;
        break;
    end
end

if isempty(target_sheet)
    target_sheet = fx_sheets{1};
    sm_idx = 3;
    warning('⚠️ 未在任一 sheet 发现 “sm” 列，将使用第一张表，并在第3列位置插入。');
end

fprintf('📄 目标工作表：%s；sm 列索引：%d\n', target_sheet, sm_idx);

C = readcell(flux_file, 'Sheet', target_sheet);
assert(size(C,1) >= 2, '❌ 目标 sheet 为空或仅表头：%s', target_sheet);

header = C(1,:);
data   = C(2:end,:);
n = size(data,1);

sma1_col = cell(n,1);
sma2_col = cell(n,1);
rb_col   = cell(n,1);
phi_col  = cell(n,1);
rs_col   = cell(n,1);

for i = 1:n
    row_flag = lower(strtrim(string(data{i, sm_idx})));

    sma1_col{i}=''; sma2_col{i}=''; rb_col{i}=''; phi_col{i}=''; rs_col{i}='';

    if ~strcmp(row_flag,'sm'), continue; end

    p = str2double(string(data{i, 1}));
    s = str2double(string(data{i, 2}));

    if ~(isfinite(p) && isfinite(s))
        sma1_col{i}='NA'; sma2_col{i}='NA';
        rb_col{i}='NA';   phi_col{i}='NA'; rs_col{i}='NA';
        continue;
    end

    k = sprintf('%g|%g', p, s);

    sma1_txt = ''; if isKey(map_sma1,k), sma1_txt = map_sma1(k); end
    sma2_txt = ''; if isKey(map_sma2,k), sma2_txt = map_sma2(k); end

    sma1_col{i} = local_str_or_NA(sma1_txt);
    sma2_col{i} = local_str_or_NA(sma2_txt);

    rb_val = NaN; if isKey(map_rb,k),  rb_val = map_rb(k);  end
    ph_val = NaN; if isKey(map_phi,k), ph_val = map_phi(k); end
    rs_val = NaN; if isKey(map_rs,k),  rs_val = map_rs(k);  end
    if ~isfinite(rs_val), rs_val = 2.65; end

    rb_col{i}  = local_num_or_NA(rb_val);
    phi_col{i} = local_num_or_NA(ph_val);
    rs_col{i}  = local_num_or_NA(rs_val);
end

% 插入 5 列到 sm 列前
new_header = [header(1:sm_idx-1), {'SM_annual_1','SM_annual_2','ρb','φ','ρs'}, header(sm_idx:end)];
new_data   = cell(n, numel(new_header));

for i = 1:n
    left_part  = data(i, 1:sm_idx-1);
    right_part = data(i, sm_idx:end);
    new_data(i,:) = [left_part, {sma1_col{i}, sma2_col{i}, rb_col{i}, phi_col{i}, rs_col{i}}, right_part];
end

%% ===================== Step 3: 仅 sm 行换算 WFPS(%)（DOY 行不动）（仅内存） =====================
fprintf('\n📌 Step 3: 仅 sm 行：统一换算到 WFPS(%%)（DOY 行不动；仅内存）...\n');

idx_SM1 = find(strcmpi(string(new_header),'SM_annual_1'),1,'first');
idx_SM2 = find(strcmpi(string(new_header),'SM_annual_2'),1,'first');
idx_rb  = find(strcmpi(string(new_header),'ρb'),1,'first');
idx_phi = find(strcmpi(string(new_header),'φ'),1,'first');
idx_rs  = find(strcmpi(string(new_header),'ρs'),1,'first');

sm_idx2 = find(strcmpi(strtrim(string(new_header)),'sm'),1,'first');
if isempty(sm_idx2), sm_idx2 = find(contains(lower(string(new_header)),'sm'),1,'first'); end

assert(all(~isnan([idx_SM1,idx_SM2,idx_rb,idx_phi,idx_rs,sm_idx2])), ...
    '❌ 缺少必要列（SM_annual_1/2/ρb/φ/ρs/sm）。');

% 数据区起始列（按原逻辑：从第11列开始）
start_col = 11;
last_col  = numel(new_header);
mcols = start_col:last_col;
p = numel(mcols);

is_sm   = strcmpi(strtrim(string(new_data(:, sm_idx2))), 'sm');
type_str = lower(strtrim(string(new_data(:, idx_SM1))));
unit_str = lower(strtrim(strrep(string(new_data(:, idx_SM2)),'％','%')));

rbv = cellfun(@local_num_coerce, new_data(:, idx_rb));
phv_pct = cellfun(@local_num_coerce, new_data(:, idx_phi));
rsv = cellfun(@local_num_coerce, new_data(:, idx_rs));
rsv(~isfinite(rsv)) = 2.65;

% phi 小数（0-1），优先用 φ 列；缺失则由 1-rb/rs 推断
phi_small = phv_pct;
mask_gt1 = phi_small>1 & ~isnan(phi_small);
phi_small(mask_gt1) = phi_small(mask_gt1)/100;

needPhi = isnan(phi_small) & isfinite(rbv) & isfinite(rsv);
phi_small(needPhi) = 1 - rbv(needPhi)./rsv(needPhi);

phi_small = max(min(phi_small,0.95),0.05);

% 读数值矩阵
S = string(new_data(:, mcols));
S = replace(S,'％','%');
S = regexprep(S,'\s+','');
S = replace(S,',','.');

maskPct = endsWith(S,'%');
S(maskPct) = extractBefore(S(maskPct), strlength(S(maskPct)));

X = str2double(S);
Xfrac = apply_ratio_mat(X, unit_str); % 将 1/%/‰ 统一为 fraction

% 将不同 type 转成 WFPS(%)
W = NaN(size(X));
rows_wfps = is_sm & (type_str=="wfps");
rows_vwc  = is_sm & (type_str=="vwc");
rows_gwc  = is_sm & (type_str=="gwc");
rows_afp  = is_sm & (type_str=="afp");

if any(rows_wfps)
    W(rows_wfps,:) = 100 .* Xfrac(rows_wfps,:);
end

if any(rows_vwc)
    ps = phi_small(:)*ones(1,p);
    ok = rows_vwc & isfinite(phi_small);
    W(ok,:) = 100 .* (Xfrac(ok,:) ./ ps(ok,:));
end

if any(rows_gwc)
    ps = phi_small(:)*ones(1,p);
    rb = rbv(:)*ones(1,p);
    ok = rows_gwc & isfinite(phi_small) & isfinite(rbv);
    W(ok,:) = 100 .* ((Xfrac(ok,:) .* rb(ok,:)) ./ ps(ok,:));
end

if any(rows_afp)
    ps = phi_small(:)*ones(1,p);
    ok = rows_afp & isfinite(phi_small);
    R = min(max(Xfrac(ok,:) ./ ps(ok,:), 0), 1);
    W(ok,:) = 100 .* (1 - R);
end

% 仅 sm 行：将单位类型改为 wfps / %
conv_mask = rows_wfps | rows_vwc | rows_gwc | rows_afp;
new_data(conv_mask, idx_SM1) = {'wfps'};
new_data(conv_mask, idx_SM2) = {'%'};

new_data_conv = new_data;
if any(is_sm)
    new_data_conv(is_sm, mcols) = num2cell(W(is_sm, :));
end

%% ===================== Step 4: 仅 sm 行>100 缩放 + 插入 WFPS_scale（仅内存） =====================
fprintf('\n📌 Step 4: 仅 sm 行：max(WFPS)>100 等比缩放 + 插入 WFPS_scale（仅内存）...\n');

vmax = nan(size(new_data,1),1);
if any(is_sm)
    vmax(is_sm) = max(W(is_sm, :), [], 2, 'omitnan');
end

scale = repmat({''}, size(new_data,1), 1);
idx_na   = is_sm & ~isfinite(vmax);
idx_over = is_sm & isfinite(vmax) & (vmax > 100 + 1e-9);
idx_ok   = is_sm & isfinite(vmax) & ~idx_over;

scale(idx_na) = {'NA'};
scale(idx_ok) = {1};

if any(idx_over)
    k = 100 ./ vmax(idx_over);
    W(idx_over, :) = W(idx_over, :) .* k;
    scale(idx_over) = num2cell(k);
end

if any(is_sm)
    new_data_conv(is_sm, mcols) = num2cell(W(is_sm, :));
end

insert_pos = 11; % 在第11列前插入 WFPS_scale
new_header_scaled = [new_header(1:insert_pos-1), {'WFPS_scale'}, new_header(insert_pos:end)];
new_data_scaled   = [new_data_conv(:,1:insert_pos-1), scale, new_data_conv(:,insert_pos:end)];

% 更新数据区列索引（因为插入了一列）
mcols_scaled = (start_col+1) : (last_col+1); % 原 start_col=11 -> 12
col_scale = find(strcmpi(string(new_header_scaled),'WFPS_scale'),1);

% 二次兜底：缩放后仍>100则再缩
if any(is_sm)
    tmpW = cellfun(@local_num_coerce, new_data_scaled(:, mcols_scaled));
    vmax2 = nan(size(new_data_scaled,1),1);
    vmax2(is_sm) = max(tmpW(is_sm,:), [], 2, 'omitnan');
    over2 = is_sm & isfinite(vmax2) & (vmax2 > 100 + 1e-9);
    if any(over2)
        idx = find(over2).';
        for t = idx
            row_vals = tmpW(t,:);
            k2 = 100 / max(row_vals, [], 'omitnan');
            row_vals = row_vals * k2;
            new_data_scaled(t, mcols_scaled) = num2cell(row_vals);

            sc_old = local_num_coerce(new_data_scaled{t, col_scale});
            if ~isfinite(sc_old), sc_old = 1; end
            new_data_scaled{t, col_scale} = sc_old * k2;
        end
    end
end

%% ===================== Step 5: TRIM（删除第 5/6/7/8/9/11 列）（仅内存） =====================
fprintf('\n📌 Step 5: TRIM（去掉第 5/6/7/8/9/11 列；仅内存）...\n');

C_scaled = [new_header_scaled; new_data_scaled];
ncol = size(C_scaled,2);

drop = [5 6 7 8 9 11];
drop = drop(drop <= ncol);

keep = setdiff(1:ncol, drop);
C_trim = C_scaled(:, keep);

%% ===================== Step 6: DOY → 月均（仅内存，已修复 DOY 列定位） =====================
fprintf('\n📌 Step 6: DOY → 月均（若存在 DOY 行；仅内存）...\n');

raw = C_trim;
[R,Cc] = size(raw);

% ✅ FIX：自动找 doy 列，而不是固定用第5列
doy_col = 0;
maxhit = 0;
scanCols = 1:min(20, Cc);
for cc = scanCols
    colv = raw(2:end, cc);
    hit = 0;
    for rr = 1:numel(colv)
        v = colv{rr};
        if (ischar(v) || isstring(v)) && strcmpi(strtrim(string(v)), "doy")
            hit = hit + 1;
        end
    end
    if hit > maxhit
        maxhit = hit;
        doy_col = cc;
    end
end
if doy_col == 0
    doy_col = 5;
end
fprintf('✅ Step6: detected DOY column index = %d (hits=%d)\n', doy_col, maxhit);

output_rows = {};
any_found = false;

for r = 2:(R-1)
    tag = raw{r,doy_col};
    if ~(ischar(tag) || isstring(tag)) || ~strcmpi(strtrim(string(tag)), 'doy')
        continue;
    end
    any_found = true;

    % 年份：优先该 DOY 行第3列；失败则第3行扫一遍
    year_calc = 2000;
    yraw = [];
    if Cc >= 3, yraw = raw{r,3}; end
    [yok, yval] = local_parse_year_scalar(yraw);
    if yok
        year_calc = yval;
    else
        if R >= 3
            for cc = 1:Cc
                [yok2, yval2] = local_parse_year_scalar(raw{3,cc});
                if yok2, year_calc = yval2; break; end
            end
        end
    end

    year_calc = round(year_calc);

    % ✅ FIX：DOY 与值从 doy_col+1 到最后，自动识别数值
    c0 = min(doy_col+1, Cc);
    if Cc < c0, continue; end

    doy_list = cellfun(@local_to_num_safe, raw(r,   c0:Cc), 'UniformOutput', false);
    val_list = cellfun(@local_to_num_safe, raw(r+1, c0:Cc), 'UniformOutput', false);

    doy = cellfun(@local_to_scalar, doy_list);
    val = cellfun(@local_to_scalar, val_list);

    valid = ~isnan(doy) & ~isnan(val);
    doy = doy(valid);
    val = val(valid);

    if isempty(doy), continue; end
    doy(doy < 1) = 1;

    months = month(datetime(year_calc, 1, 1) + days(doy - 1));
    mon_vals = nan(1,12);

    for m = 1:12
        idx = (months == m);
        if any(idx), mon_vals(m) = mean(val(idx)); end
    end

    % ✅ 统一输出 monout_cell 结构：前4列 paper/study/year/layer，col5=mon，col6..17=1..12
    row1 = cell(1, 5 + 12);
    row2 = cell(1, 5 + 12);

    for cc = 1:4
        if cc <= Cc
            row1{cc} = raw{r,cc};
            row2{cc} = raw{r+1,cc};
        else
            row1{cc} = [];
            row2{cc} = [];
        end
    end

    row1{5} = 'mon';
    row2{5} = 'sm';

    row1(6:17) = num2cell(1:12);
    row2(6:17) = num2cell(mon_vals);

    output_rows(end+1:end+2, 1) = {row1; row2}; %#ok<AGROW>
end

if ~any_found
    monout_cell = {'NO_DOY'};
    fprintf('⚠️ 未找到 DOY 行，monout_cell = NO_DOY\n');
else
    monout_cell = vertcat(output_rows{:});
    fprintf('✅ 已在内存生成 monout_cell（%d 行）\n', size(monout_cell,1));
end

%% ===================== Step 7: 制作 move 模板（READ move_template_SM.xlsx，仅内存） =====================
fprintf('\n📌 Step 7: 读取 move_template_SM.xlsx 作为列结构（不写回，只在内存对齐 papers/studies）...\n');

tpl = readtable(tpl_file, 'Sheet', picked_sheet, 'VariableNamingRule','preserve','UseExcel',false);
papers  = str2double(regexprep(string(tpl.("Paper_number")),'[^\d\.\-]',''));
studies = str2double(regexprep(string(tpl.("Study_number")),'[^\d\.\-]',''));
target_h = numel(papers);

T0 = readtable(move_file, 'VariableNamingRule','preserve', 'UseExcel', false);
vnames0 = T0.Properties.VariableNames;
h0 = height(T0);

Tmove = table();
for v = 1:numel(vnames0)
    vn = vnames0{v};
    col = T0.(vn);

    if isvector(col)
        w = 1; col = col(:);
    else
        w = size(col,2);
    end

    if isnumeric(col) || islogical(col)
        Tmove.(vn) = NaN(target_h, w);
    elseif isdatetime(col)
        Tmove.(vn) = repmat(NaT, target_h, w);
    elseif isduration(col)
        Tmove.(vn) = repmat(seconds(NaN), target_h, w);
    elseif isstring(col)
        Tmove.(vn) = repmat(missing, target_h, w);
    elseif iscategorical(col)
        if h0 > 0
            tmp = repmat(col(1,:), target_h, 1);
            tmp(:) = categorical(missing);
            Tmove.(vn) = tmp;
        else
            Tmove.(vn) = categorical(strings(target_h, w));
        end
    elseif iscell(col)
        tmp = cell(target_h, w);
        tmp(:) = {[]};
        Tmove.(vn) = tmp;
    else
        s = string(col);
        Tmove.(vn) = repmat(missing, target_h, size(s,2));
    end
end

h_copy = min(h0, target_h);
for v = 1:numel(vnames0)
    vn = vnames0{v};
    col_src = T0.(vn);
    col_dst = Tmove.(vn);

    if isvector(col_src), col_src = col_src(:); end

    if iscategorical(col_dst) && ~iscategorical(col_src) && h_copy > 0
        try
            col_src = categorical(col_src);
        catch
            col_src = categorical(string(col_src));
        end
    end

    if isnumeric(col_dst) && ~isa(col_src,'double')
        col_src = double(col_src);
    end

    [~, w_dst] = size(col_dst);
    [~, w_src] = size(col_src);
    w_copy = min(w_dst, w_src);

    if h_copy > 0 && w_copy > 0
        col_dst(1:h_copy, 1:w_copy) = col_src(1:h_copy, 1:w_copy);
    end
    Tmove.(vn) = col_dst;
end

mvNames = string(Tmove.Properties.VariableNames);
if numel(mvNames) < 16
    error('❌ move_template_SM.xlsx 列数不足，至少需要 16 列（papers,studies,year,layer,m1..m12）。');
end

mvNames(1:4) = ["papers","studies","year","layer"];
Tmove.Properties.VariableNames = cellstr(mvNames);

Tmove.papers  = papers(:);
Tmove.studies = studies(:);

fprintf('✅ Step7: 已在内存构造对齐 move 表：原=%d 行，现=%d 行（move_template_SM.xlsx 不改动）\n', h0, height(Tmove));

%% ===================== Step 8: yearly-monthly 展开 + 回填主行（仅内存） =====================
fprintf('\n📌 Step 8: 构造 yearly-monthly 展开 + 回填主行（月均；仅内存）...\n');

target_w = width(Tmove);
header_move = Tmove.Properties.VariableNames;
MCOLS = 5:16;

mon_rows = {};
mon_ord  = [];
ord_ctr  = 0;

if iscell(monout_cell) && size(monout_cell,1) >= 2 && size(monout_cell,2) >= 6 && ~isequal(monout_cell{1,1},'NO_DOY')
    [MR, ~] = size(monout_cell);
    for r = 1:2:(MR-1)
        meta = monout_cell(r, :);
        vals = monout_cell(r+1, :);

        needCols = 17;
        if size(meta,2) < needCols, meta(1,end+1:needCols) = {[]}; end
        if size(vals,2) < needCols, vals(1,end+1:needCols) = {[]}; end

        paper = meta{1,1};
        study = meta{1,2};
        year  = meta{1,3};
        layer = meta{1,4};
        if isempty(paper) || isempty(study), continue; end

        mvals = cell(1,12);
        for k = 1:12
            mvals{k} = vals{1, 5 + k};
        end

        row16 = cell(1,16);
        row16(1:4) = {paper, study, year, layer};
        row16(5:16) = mvals;

        ord_ctr = ord_ctr + 1;
        mon_rows(end+1,1:16) = row16; %#ok<AGROW>
        mon_ord(end+1,1) = ord_ctr; %#ok<AGROW>
    end
end

if isempty(mon_rows)
    final_expanded_cell = [header_move; table2cell(Tmove)];
else
    colnamesMon = ["paper","study","year","layer", ...
        "m1","m2","m3","m4","m5","m6","m7","m8","m9","m10","m11","m12"];

    Tm = cell2table(mon_rows, 'VariableNames', cellstr(colnamesMon));
    Tm.ord  = mon_ord(:);

    Tm.paper = local_numify_col(Tm.paper);
    Tm.study = local_numify_col(Tm.study);
    Tm.year  = local_numify_col(Tm.year);
    for ii = 5:16
        vn = colnamesMon(ii);
        Tm.(vn) = local_numify_col(Tm.(vn));
    end

    final_rows = {};
    for i = 1:height(Tmove)
        paper_i = Tmove.papers(i);
        study_i = Tmove.studies(i);

        mask = (Tm.paper == paper_i) & (Tm.study == study_i);
        Tmatch = Tm(mask, :);

        if ~isempty(Tmatch)
            Tmatch = sortrows(Tmatch, {'year','ord'});
            for j = 1:height(Tmatch)
                out = repmat({[]}, 1, target_w);
                out{1} = Tmatch.paper(j);
                out{2} = Tmatch.study(j);
                out{3} = Tmatch.year(j);
                out{4} = Tmatch.layer(j);
                for m = 1:12
                    out{MCOLS(m)} = Tmatch{j, sprintf('m%d',m)};
                end
                final_rows(end+1,:) = out; %#ok<AGROW>
            end
        end

        final_rows(end+1,:) = table2cell(Tmove(i,:)); %#ok<AGROW>
    end

    T_year = groupsummary(Tm, {'paper','study','year'}, 'mean', cellstr(colnamesMon(5:end)));
    for m = 1:12
        src = sprintf('mean_m%d',m);
        dst = sprintf('m%d',m);
        T_year.(dst) = T_year.(src);
        T_year.(src) = [];
    end

    varList = arrayfun(@(mm)sprintf('m%d',mm), 1:12, 'UniformOutput', false);
    T_main = groupsummary(T_year, {'paper','study'}, 'mean', varList);
    for m = 1:12
        src = sprintf('mean_m%d',m);
        dst = sprintf('m%d',m);
        T_main.(dst) = T_main.(src);
        T_main.(src) = [];
    end

    for k = 1:size(final_rows,1)
        pk = local_numify(final_rows{k,1});
        sk = local_numify(final_rows{k,2});
        yk = local_numify(final_rows{k,3});
        if isnan(pk) || isnan(sk), continue; end
        if isnan(yk)
            idx = find(T_main.paper == pk & T_main.study == sk, 1, 'first');
            if ~isempty(idx)
                for m = 1:12
                    final_rows{k, MCOLS(m)} = { T_main{idx, sprintf('m%d',m)} };
                end
            end
        end
    end

    numeric_idx = [1 2 3 MCOLS];
    for rr = 1:size(final_rows,1)
        for cc = 1:target_w
            v = local_unwrap_scalar_cell(final_rows{rr,cc});
            if (isa(v,'missing')) || (isstring(v) && isscalar(v) && ismissing(v))
                if any(cc == numeric_idx), v = NaN; else, v = ''; end
            end
            final_rows{rr,cc} = v;
        end
    end

    final_expanded_cell = [header_move; final_rows];
end

%% ===================== Step 9: 从主行生成 M_main（仅内存） =====================
fprintf('\n📌 Step 9: 从 expanded 主行生成 M_main（仅内存）...\n');

hdr = string(final_expanded_cell(1,:));
data_exp = final_expanded_cell(2:end,:);

col_p = find(strcmpi(hdr,'papers'),1); if isempty(col_p), col_p=1; end
col_s = find(strcmpi(hdr,'studies'),1); if isempty(col_s), col_s=2; end
col_y = find(strcmpi(hdr,'year'),1);   if isempty(col_y), col_y=3; end

mcols = zeros(1,12);
for m = 1:12
    hit = find(strcmpi(hdr, sprintf('m%d',m)), 1);
    if ~isempty(hit), mcols(m) = hit; end
end
if any(mcols==0)
    mcols = 5:16;
end

mkKey = @(p,s) sprintf('%.15g|%.15g', p, s);

M_main = containers.Map('KeyType','char','ValueType','any');
main_cnt = 0;

for i = 1:size(data_exp,1)
    p = local_numify(data_exp{i,col_p});
    s = local_numify(data_exp{i,col_s});
    y = local_numify(data_exp{i,col_y});

    if ~(isfinite(p) && isfinite(s)), continue; end
    if isfinite(y), continue; end

    vals12 = NaN(1,12);
    for m = 1:12
        vals12(m) = local_f4_numify(data_exp{i, mcols(m)});
    end

    M_main(mkKey(p,s)) = vals12;
    main_cnt = main_cnt + 1;
end

fprintf('✅ M_main 已生成：%d 条主行键值\n', main_cnt);

%% ===================== Step 10: 回填到 Studies and Fluxes.xlsx（唯一写盘，IN-PLACE，已修复） =====================
fprintf('\n📌 Step 10: 将 Step9 主行月均回填到 Studies and Fluxes.xlsx（IN-PLACE，只写 12 列，不覆盖整表）...\n');

% ---- find correct sheet (robust) ----
try
    tarSheets = sheetnames(target_flux);
catch
    [~, tarSheets] = xlsfinfo(target_flux);
end
if isempty(tarSheets), error('❌ target_flux 无可读 sheet'); end

pickedTarSheet = '';
Ttar = table();

for ss = 1:numel(tarSheets)
    Ttmp = readtable(target_flux, 'Sheet', tarSheets{ss}, 'VariableNamingRule','preserve', 'UseExcel', false);
    vtmp = string(Ttmp.Properties.VariableNames);
    if any(strcmpi(vtmp,'Paper_number')) && any(strcmpi(vtmp,'Study_number'))
        pickedTarSheet = tarSheets{ss};
        Ttar = Ttmp;
        break;
    end
end

if isempty(pickedTarSheet)
    pickedTarSheet = tarSheets{1};
    Ttar = readtable(target_flux, 'Sheet', pickedTarSheet, 'VariableNamingRule','preserve', 'UseExcel', false);
end
fprintf('✅ Step10: target sheet used: %s\n', pickedTarSheet);

vars_tar = string(Ttar.Properties.VariableNames);

% key 列（目标文件必须有）
colPname = local_find_col(vars_tar, {'Paper_number'});
colSname = local_find_col(vars_tar, {'Study_number'});
if isempty(colPname) || isempty(colSname)
    error('❌ Step10：目标文件中找不到 Paper_number / Study_number。');
end

paper_tar = str2double(regexprep(string(Ttar.(colPname)),'[^\d\.\-]',''));
study_tar = str2double(regexprep(string(Ttar.(colSname)),'[^\d\.\-]',''));

% ✅ 写回列优先级：SM_m1..SM_m12 -> m1..m12 -> Jan..Dec
useCols = strings(1,12);
modeName = '';

% 1) SM_m1..SM_m12
sm_cols = strings(1,12);
for m = 1:12
    hit = local_find_col(vars_tar, {sprintf('SM_m%d',m)});
    if ~isempty(hit), sm_cols(m) = hit; end
end
if all(sm_cols~="")
    useCols = sm_cols;
    modeName = 'SM_m1..SM_m12';
else
    % 2) m1..m12
    m_cols = strings(1,12);
    for m = 1:12
        hit = local_find_col(vars_tar, {sprintf('m%d',m)});
        if ~isempty(hit), m_cols(m) = hit; end
    end
    if all(m_cols~="")
        useCols = m_cols;
        modeName = 'm1..m12';
    else
        % 3) Jan..Dec
        janDec = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
        jd_cols = strings(1,12);
        for m = 1:12
            hit = local_find_col(vars_tar, {char(janDec(m))});
            if ~isempty(hit), jd_cols(m) = hit; end
        end
        if all(jd_cols~="")
            useCols = jd_cols;
            modeName = 'Jan..Dec';
        else
            error('❌ 目标文件未找到完整 SM_m1..SM_m12 / m1..m12 / Jan..Dec（不允许新增列，只覆盖）。');
        end
    end
end

fprintf('✅ Step10: overwrite mode = %s\n', modeName);

% ---- build existing numeric block (preserve rows w/o new values) ----
nRowTar = height(Ttar);
oldBlk = NaN(nRowTar, 12);

for m = 1:12
    vn = char(useCols(m));
    col = Ttar.(vn);

    if isnumeric(col)
        oldBlk(:,m) = double(col);
    else
        tmp = NaN(nRowTar,1);
        for i = 1:nRowTar
            tmp(i) = local_f4_numify(col(i));
        end
        oldBlk(:,m) = tmp;
    end
end

newBlk = oldBlk;

% ---- overwrite only matched keys and finite values ----
filled_rows = 0;
filled_cells = 0;

for i = 1:nRowTar
    p = paper_tar(i);
    s = study_tar(i);
    if ~(isfinite(p) && isfinite(s)), continue; end

    k = mkKey(p,s);
    if isKey(M_main, k)
        vals12 = double(M_main(k));
        didRow = false;
        for m = 1:12
            if ~isnan(vals12(m))
                newBlk(i,m) = vals12(m);
                filled_cells = filled_cells + 1;
                didRow = true;
            end
        end
        if didRow
            filled_rows = filled_rows + 1;
        end
    end
end

% ---- write back ONLY these 12 columns (NO writetable full overwrite) ----
% Column indices (Excel position) from vars_tar
idxCols = zeros(1,12);
for m = 1:12
    idxCols(m) = find(strcmp(vars_tar, useCols(m)), 1, 'first');
end

% 如果 12 列连续，就一次写；否则逐列写（更安全）
isConsecutive = all(diff(idxCols) == 1);

if isConsecutive
    startCol = idxCols(1);
    startCell = sprintf('%s2', local_colnum2excel(startCol));

    Cblk = num2cell(newBlk);
    Cblk(isnan(newBlk)) = {[]};

    writecell(Cblk, target_flux, 'Sheet', pickedTarSheet, 'Range', startCell);
else
    for m = 1:12
        startCell = sprintf('%s2', local_colnum2excel(idxCols(m)));
        Ccol = num2cell(newBlk(:,m));
        Ccol(isnan(newBlk(:,m))) = {[]};
        writecell(Ccol, target_flux, 'Sheet', pickedTarSheet, 'Range', startCell);
    end
end

fprintf('✅ Step10 done (PARTIAL overwrite ONLY 12 cols): %s\n', target_flux);
fprintf(' Matched & filled rows: %d\n', filled_rows);
fprintf(' Cells overwritten (non-NaN only): %d\n', filled_cells);
fprintf(' Column mode used: %s\n', modeName);

%% ======================= 工具函数区（统一放末尾） =======================

function y = local_phi_to_pct(v, unitStr)
    if isnan(v), y = NaN; return; end
    u = lower(strtrim(string(unitStr)));

    if any(u == ["%","percent","百分比"])
        y = v;
    elseif any(u == ["v/v","cm3/cm3","m3/m3","m³/m³","fraction","小数","比值"])
        y = v * 100;
    elseif u == "" || u == "na" || u == "nan" || u == "unknown"
        if v <= 1.5
            y = v * 100;
        elseif v <= 100
            y = v;
        else
            y = NaN;
        end
    else
        if v <= 1.5
            y = v * 100;
        elseif v <= 100
            y = v;
        else
            y = NaN;
        end
    end

    if ~(y >= 0 && y <= 100), y = NaN; end
end

function out = local_num_or_NA(x)
    if isnan(x), out = 'NA'; else, out = x; end
end

function out = local_str_or_NA(x)
    xs = string(x);
    if strlength(xs)==0 || ismissing(xs)
        out = 'NA';
    else
        out = char(xs);
    end
end

function v = local_num_coerce(x)
    xs = string(x);
    if ismissing(xs) || strlength(strtrim(xs))==0
        v = NaN; return;
    end
    xs = strrep(xs,'％','%');
    xs = regexprep(xs,'\s+','');
    xs = strrep(xs, ',', '.');
    if endsWith(xs,'%')
        xs = extractBefore(xs, strlength(xs));
    end
    v = str2double(xs);
end

function M_frac = apply_ratio_mat(M_raw, unit_str_row)
    rt = lower(strip(string(unit_str_row)));
    rt(ismissing(rt) | rt=="") = "1";

    M_frac = M_raw;

    mask1 = rt=="1" | rt=="none";
    maskp = rt=="%";
    maskm = rt=="‰" | rt=="permille" | rt=="per mille" | rt=="per_mille" | rt==string(char(8240));

    if any(maskp), M_frac(maskp,:) = M_raw(maskp,:) ./ 100; end
    if any(maskm), M_frac(maskm,:) = M_raw(maskm,:) ./ 1000; end

    maskU = ~(mask1 | maskp | maskm);
    if any(maskU)
        X = M_raw(maskU,:);
        X(X>1 & X<=100)  = X(X>1 & X<=100)/100;
        X(X>100 & X<=1000) = X(X>100 & X<=1000)/1000;
        M_frac(maskU,:) = X;
    end
end

function y = local_to_num_safe(x)
    if isempty(x)
        y = NaN;
    elseif isnumeric(x) || islogical(x)
        y = double(x);
    elseif ischar(x) || isstring(x)
        y = str2double(string(x));
    else
        try
            y = str2double(string(x));
            if numel(y) ~= 1 || isnan(y), y = NaN; end
        catch
            y = NaN;
        end
    end
end

function s = local_to_scalar(v)
    if isnumeric(v) || islogical(v)
        if isempty(v), s = NaN; else, s = double(v(1)); end
    else
        s = NaN;
    end
end

function [ok, yy] = local_parse_year_scalar(x)
    yy = NaN; ok = true;
    if isempty(x)
        ok = false;
    elseif isnumeric(x) && isfinite(x)
        yy = x;
    elseif ischar(x) || isstring(x)
        ystr = regexp(char(x), '\d{4}', 'match');
        if isempty(ystr), ok=false; else, yy = str2double(ystr{1}); end
    else
        ok = false;
    end
end

function z = local_unwrap_scalar_cell(z)
    while iscell(z) && isscalar(z)
        z = z{1};
    end
    if isstring(z) && isscalar(z) && ~ismissing(z)
        z = char(z);
    end
end

function v = local_numify_col(v)
    if iscell(v)
        v = cellfun(@local_numify, v);
    elseif isnumeric(v)
        v = double(v);
    elseif isstring(v) || ischar(v) || iscategorical(v)
        v = str2double(string(v));
    else
        v = arrayfun(@local_numify, num2cell(v));
    end
end

function y = local_numify(x)
    if iscell(x)
        if isempty(x), y = NaN;
        elseif isscalar(x), y = local_numify(x{1});
        else, y = NaN;
        end
    elseif isnumeric(x)
        if isempty(x), y = NaN;
        elseif isscalar(x), y = double(x);
        else, y = NaN;
        end
    elseif isstring(x) || ischar(x)
        xs = strtrim(string(x));
        if isscalar(xs), y = str2double(xs); else, y = NaN; end
    elseif isempty(x)
        y = NaN;
    else
        y = NaN;
    end
end

function y = local_f4_numify(x)
    if iscell(x)
        if isempty(x), y = NaN;
        elseif isscalar(x), y = local_f4_numify(x{1});
        else, y = NaN;
        end
    elseif isnumeric(x)
        if isempty(x), y = NaN;
        elseif isscalar(x), y = double(x);
        else, y = NaN;
        end
    elseif isstring(x) || ischar(x)
        xs = strtrim(string(x));
        if isscalar(xs), y = str2double(xs); else, y = NaN; end
    elseif isempty(x)
        y = NaN;
    else
        y = NaN;
    end
end

function colname = local_find_col(vars_all, candidates)
    colname = '';
    vlow = lower(strtrim(vars_all));
    for ii = 1:numel(candidates)
        cand = lower(strtrim(string(candidates{ii})));
        hit = find(vlow == cand, 1, 'first');
        if ~isempty(hit)
            colname = char(vars_all(hit));
            return;
        end
    end
end

function letters = local_colnum2excel(n)
    letters = '';
    while n > 0
        r = mod(n-1, 26);
        letters = [char(r + 'A') letters];
        n = floor((n-1) / 26);
    end
end
