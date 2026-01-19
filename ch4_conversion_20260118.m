clear;

%% =========================================================
%  ALL inputs & outputs are restricted to ONE folder:
%  D:\Users\jiaweiChiang\Desktop\Supplymentary_code
%
%  ✅ move_template_CH4.xlsx: NOT changed (kept)  —— 不写回、不覆盖
%  ✅ NEW template: Studies and Fluxes-unit.xlsx (auto-detect sheet)
%  ✅ Step10 target: Studies and Fluxes.xlsx (PARTIAL overwrite only 17 cols)
%  ✅ m1–m12 insertion: between Monthly_9 and Seasonal   (filtrate_* columns removed)
%  ✅ Only Step10 writes back to Excel (no Final17 excel export)
%% =========================================================
ROOT = 'D:\Users\jiaweiChiang\Desktop\Supplymentary_code';
cd(ROOT);

% ------------------ Inputs (must exist in ROOT) ------------------
RAW_XLSX  = fullfile(ROOT, 'CH4 uptake data_CH4 FLUX_1_1410.xlsx');   % raw DOY daily file (source)
MOVE_XLSX = fullfile(ROOT, 'move_template_CH4.xlsx');                 % move template (DO NOT change)

% ------------------ NEW template file (flux-unit) ------------------
TEMPLATE_XLSX = fullfile(ROOT, 'Studies and Fluxes-unit.xlsx');
assert(isfile(TEMPLATE_XLSX), '❌ Cannot find TEMPLATE_XLSX: %s', TEMPLATE_XLSX);

% ------------------ Step10 target (flux) ------------------
TARGET_FLUX_XLSX = fullfile(ROOT, 'Studies and Fluxes.xlsx');
assert(isfile(TARGET_FLUX_XLSX), '❌ Cannot find target file: %s', TARGET_FLUX_XLSX);

% ------------------ optional backup switch ------------------
DO_BACKUP = false;  % ✅ 默认 false：严格满足“只写回 Step10”要求；需要备份就改成 true

fprintf('✅ ROOT: %s\n', ROOT);
fprintf('✅ RAW_XLSX: %s\n', RAW_XLSX);
fprintf('✅ MOVE_XLSX (kept, no overwrite): %s\n', MOVE_XLSX);
fprintf('✅ TEMPLATE_XLSX (flux-unit): %s\n', TEMPLATE_XLSX);
fprintf('✅ TARGET_FLUX_XLSX (partial overwrite only 17 cols): %s\n', TARGET_FLUX_XLSX);

%% ===========================
% Step 1 (one-sentence summary):
% Aggregate raw DOY-based daily flux data into monthly mean values (m1–m12) per record.
%% ===========================

assert(isfile(RAW_XLSX), '❌ Cannot find RAW_XLSX: %s', RAW_XLSX);

try
    sheets = sheetnames(RAW_XLSX);
catch
    [~, sheets] = xlsfinfo(RAW_XLSX);
end
if isempty(sheets)
    error('❌ The RAW workbook has no readable sheets.');
end

skipped_info = {};  % {Sheet, RowIndex, ColIndex, Reason}

% Only process the first sheet
for s = 1
    sheet = sheets{s};

    try
        [~, ~, raw] = xlsread(RAW_XLSX, sheet);
    catch
        warning(['⚠️ Failed to read sheet: ' sheet]);
        skipped_info(end+1,:) = {sheet, [], [], 'sheet_read_failed'}; %#ok<AGROW>
        continue;
    end
    if isempty(raw)
        warning(['⚠️ Empty sheet: ' sheet]);
        skipped_info(end+1,:) = {sheet, [], [], 'sheet_empty'}; %#ok<AGROW>
        continue;
    end

    [R,C] = size(raw);
    output_rows = {};
    any_found = false;

    for r = 1:R-1
        tag = [];
        if C >= 5, tag = raw{r,5}; end
        if ~(ischar(tag) || isstring(tag)) || ~strcmpi(strtrim(char(tag)), 'doy')
            continue;
        end
        any_found = true;

        % ---- infer year ----
        year_calc = 2000;
        yraw = []; if C >= 3, yraw = raw{r,3}; end
        [yok, yval] = parse_year_scalar(yraw);
        if yok
            year_calc = yval;
        else
            if R >= 3
                for cc = 1:C
                    [yok2, yval2] = parse_year_scalar(raw{3,cc});
                    if yok2, year_calc = yval2; break; end
                end
            end
        end

        c0 = 6;
        if C < c0
            skipped_info(end+1,:) = {sheet, r, 5, 'no_numeric_after_c5'}; %#ok<AGROW>
            continue;
        end

        doy_list = cellfun(@to_num_safe, raw(r,   c0:C), 'UniformOutput', false);
        val_list = cellfun(@to_num_safe, raw(r+1, c0:C), 'UniformOutput', false);
        doy = cellfun(@to_scalar, doy_list);
        val = cellfun(@to_scalar, val_list);

        valid = ~isnan(doy) & ~isnan(val);
        doy = doy(valid); val = val(valid);
        if isempty(doy)
            skipped_info(end+1,:) = {sheet, r, 5, 'empty_series_after_clean'}; %#ok<AGROW>
            continue;
        end
        doy(doy < 1) = 1;

        months = month(datetime(year_calc,1,1) + days(doy - 1));
        mon_vals = nan(1,12);
        for m = 1:12
            idx = (months == m);
            if any(idx)
                mon_vals(m) = mean(val(idx));
            end
        end

        % ---- output: two-row block ----
        row1 = cell(1, 5 + 12);
        row2 = cell(1, 5 + 12);
        for cc = 1:5
            if cc <= C
                row1{cc} = raw{r,cc};
                row2{cc} = raw{r+1,cc};
            else
                row1{cc} = [];
                row2{cc} = [];
            end
        end
        row1{5} = 'mon';
        row1(6:17) = num2cell(1:12);
        row2(6:17) = num2cell(mon_vals);

        output_rows(end+1:end+2, 1) = {row1; row2}; %#ok<AGROW>
    end

    if isempty(output_rows) || ~any_found
        warning(['⚠️ No valid DOY markers in sheet: ' sheet]);
        skipped_info(end+1,:) = {sheet, [], [], 'no_doy_markers'}; %#ok<AGROW>
        continue;
    end

    monout_inmem = vertcat(output_rows{:}); %#ok<NASGU>
    fprintf('✅ Step1: monout in memory created (sheet=%s, rows=%d)\n', sheet, size(monout_inmem,1));
end

assert(exist('monout_inmem','var')==1, '❌ Step1 produced no in-memory monout. Check RAW input.');

if ~isempty(skipped_info)
    fprintf('⚠️ Step1 skipped records: %d (not exported)\n', size(skipped_info,1));
end

%% ===========================
% Step 2 (one-sentence summary):
% Force the move table height and papers/studies indices to match the master template numbering (IN MEMORY ONLY).
%% ===========================

disp('⏳ Step2: aligning move table papers/studies with flux-unit numbering (IN MEMORY ONLY) ...');

assert(isfile(MOVE_XLSX), '❌ Cannot find MOVE_XLSX: %s', MOVE_XLSX);

% ---- read flux-unit: auto-detect correct sheet ----
needTplCols = ["Paper_number","Study_number"];
[tpl_sheet, template_tbl] = read_sheet_with_cols(TEMPLATE_XLSX, needTplCols);
fprintf('✅ Step2: flux-unit sheet used: %s\n', tpl_sheet);

papers  = str2double(regexprep(string(template_tbl.Paper_number),'[^\d\.\-]',''));
studies = str2double(regexprep(string(template_tbl.Study_number),'[^\d\.\-]',''));
target_h = numel(papers);

% ---- read move template (NO overwrite) ----
T0 = readtable(MOVE_XLSX,'VariableNamingRule','preserve','UseExcel',false);
h0 = height(T0);
vnames = T0.Properties.VariableNames;

% build aligned move table in memory
T2 = table();
for v = 1:numel(vnames)
    vn  = vnames{v};
    col = T0.(vn);

    if isvector(col), w = 1; col = col(:);
    else, w = size(col,2); end

    if isnumeric(col) || islogical(col)
        T2.(vn) = NaN(target_h, w);
    elseif isdatetime(col)
        T2.(vn) = repmat(NaT, target_h, w);
    elseif isduration(col)
        T2.(vn) = repmat(seconds(NaN), target_h, w);
    elseif isstring(col)
        T2.(vn) = repmat(missing, target_h, w);
    elseif iscategorical(col)
        if h0 > 0
            T2col = repmat(col(1,:), target_h, 1);
            T2col(:) = categorical(missing);
            T2.(vn) = T2col;
        else
            T2.(vn) = categorical(strings(target_h, w));
        end
    elseif iscell(col)
        tmp = cell(target_h, w); tmp(:) = {[]};
        T2.(vn) = tmp;
    else
        s = string(col);
        T2.(vn) = repmat(missing, target_h, size(s,2));
    end
end

h_copy = min(h0, target_h);
for v = 1:numel(vnames)
    vn  = vnames{v};
    col_src = T0.(vn);
    col_dst = T2.(vn);

    if isvector(col_src), col_src = col_src(:); end

    if iscategorical(col_dst) && ~iscategorical(col_src) && h_copy>0
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
    T2.(vn) = col_dst;
end

if ~ismember('papers',  vnames), T2.papers  = NaN(target_h,1); end
if ~ismember('studies', vnames), T2.studies = NaN(target_h,1); end
T2.papers  = papers(:);
T2.studies = studies(:);

fprintf('✅ Step2: move aligned in memory (old=%d, new=%d). MOVE file NOT modified.\n', h0, height(T2));

%% ===========================
% Step 3 (one-sentence summary):
% Expand the move table by inserting year/layer rows and backfill main-row monthly values using within-paper/study averages.
%% ===========================

disp('⏳ Step3: expanding move table using in-memory monout ...');

mon_raw = monout_inmem;
[MR, ~] = size(mon_raw);

Tmove = T2;  % ✅ use aligned-in-memory move
mvNames = Tmove.Properties.VariableNames;
if numel(mvNames) < 16
    error('❌ MOVE template must contain at least 16 columns (papers, studies, year, layer, m1..m12).');
end
mvNames(1:4) = {'papers','studies','year','layer'};
Tmove.Properties.VariableNames = mvNames;

target_w = width(Tmove);
header   = Tmove.Properties.VariableNames;
MCOLS    = 5:16;

mon_rows = {};
mon_ord  = [];
ord_ctr  = 0;

for r = 1:2:(MR-1)
    meta = mon_raw(r,   :);
    vals = mon_raw(r+1, :);

    needCols = 17;
    if size(meta,2) < needCols, meta(1, end+1:needCols) = {[]}; end
    if size(vals,2) < needCols, vals(1, end+1:needCols) = {[]}; end

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
    row16(1:4)  = {paper, study, year, layer};
    row16(5:16) = mvals;

    ord_ctr = ord_ctr + 1;
    mon_rows(end+1,1:16) = row16; %#ok<AGROW>
    mon_ord(end+1,1)     = ord_ctr; %#ok<AGROW>
end

colnamesMon = ["paper","study","year","layer", ...
               "m1","m2","m3","m4","m5","m6","m7","m8","m9","m10","m11","m12"];
T = cell2table(mon_rows, 'VariableNames', cellstr(colnamesMon));
T.ord = mon_ord(:);

T.paper = numify_col(T.paper);
T.study = numify_col(T.study);
T.year  = numify_col(T.year);
for ii = 5:16
    vn = colnamesMon(ii);
    T.(vn) = numify_col(T.(vn));
end

final_rows = {};
for i = 1:height(Tmove)
    paper_i = Tmove.papers(i);
    study_i = Tmove.studies(i);

    mask = (T.paper == paper_i) & (T.study == study_i);
    Tmatch = T(mask, :);

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
            final_rows(end+1, :) = out; %#ok<AGROW>
        end
    end

    final_rows(end+1, :) = table2cell(Tmove(i, :)); %#ok<AGROW>
end

% compute main means and fill main rows (year NaN)
if ~isempty(T)
    T_year = groupsummary(T, {'paper','study','year'}, 'mean', cellstr(colnamesMon(5:end)));
    for m = 1:12
        src = sprintf('mean_m%d',m); dst = sprintf('m%d',m);
        T_year.(dst) = T_year.(src); T_year.(src) = [];
    end

    varList = arrayfun(@(mm)sprintf('m%d',mm), 1:12, 'UniformOutput', false);
    T_main  = groupsummary(T_year, {'paper','study'}, 'mean', varList);
    for m = 1:12
        src = sprintf('mean_m%d',m); dst = sprintf('m%d',m);
        T_main.(dst) = T_main.(src); T_main.(src) = [];
    end

    for k = 1:size(final_rows,1)
        pk = numify(final_rows{k,1});
        sk = numify(final_rows{k,2});
        yk = numify(final_rows{k,3});

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
end

% sanitize scalars
numeric_idx = [1 2 3 MCOLS];
for rr = 1:size(final_rows,1)
    for cc = 1:target_w
        v = final_rows{rr,cc};
        v = unwrap_scalar_cell(v);
        if (isa(v,'missing')) || (isstring(v) && isscalar(v) && ismissing(v))
            if any(cc == numeric_idx), v = NaN; else, v = ''; end
        end
        final_rows{rr,cc} = v;
    end
end

if isstring(header); header = cellstr(header); end
final_out = [header; final_rows];
src_tbl_step3 = cell2table(final_out(2:end,:), 'VariableNames', final_out(1,:)); %#ok<NASGU>

%% ===========================
% Step 4 (one-sentence summary):
% Fill the aligned move template main rows (by papers/studies) using the computed monthly means (IN MEMORY ONLY).
%% ===========================

disp('⏳ Step4: filling main rows back to move template (IN MEMORY ONLY) ...');

src_tbl = src_tbl_step3;
tgt_tbl = Tmove;  % ✅ in-memory move (aligned)

if width(tgt_tbl) < 16
    error('❌ MOVE template must have at least 16 columns (meta + m1..m12).');
end
MCOLS = 5:16;

src_tbl.papers  = f4_numify_col(src_tbl.papers);
src_tbl.studies = f4_numify_col(src_tbl.studies);
tgt_tbl.papers  = f4_numify_col(tgt_tbl.papers);
tgt_tbl.studies = f4_numify_col(tgt_tbl.studies);

hasYear = ismember('year', src_tbl.Properties.VariableNames);
if hasYear
    ycol = f4_numify_col(src_tbl.year);
else
    ycol = NaN(height(src_tbl),1);
end

is_main_row = isnan(ycol);
src_main = src_tbl(is_main_row, :);

src_main{:, MCOLS} = cellfun(@f4_numify, table2cell(src_main(:, MCOLS)));

mkKey = @(p,s) sprintf('%.15g|%.15g', p, s);
keys   = strings(height(src_main),1);
values = cell(height(src_main),1);
for i = 1:height(src_main)
    keys(i)   = mkKey(src_main.papers(i), src_main.studies(i));
    values{i} = src_main{i, MCOLS};
end

M = containers.Map;
for i = 1:numel(keys)
    M(char(keys(i))) = values{i};
end

filled_cnt = 0;
for i = 1:height(tgt_tbl)
    k = mkKey(tgt_tbl.papers(i), tgt_tbl.studies(i));
    if isKey(M, k)
        tgt_tbl{i, MCOLS} = M(k);
        filled_cnt = filled_cnt + 1;
    end
end

fprintf('✅ Step4: filled %d main rows into in-memory move table. MOVE file NOT modified.\n', filled_cnt);

%% ===========================
% Step 5 (one-sentence summary):
% Insert m1–m12 into flux-unit template between Monthly_9 and Seasonal (IN MEMORY).
%% ===========================

disp('⏳ Step5: inserting m1–m12 into flux-unit template (IN MEMORY) ...');

data_tbl     = tgt_tbl;       % monthly source from move (m1..m12 in col 5..16)
% template_tbl already loaded in Step2

month_data = data_tbl(:, 5:16);
month_data.Properties.VariableNames = compose("m%d", 1:12);

n_template = height(template_tbl);
n_month    = height(month_data);

if n_month < n_template
    pad = array2table(NaN(n_template - n_month, 12), 'VariableNames', month_data.Properties.VariableNames);
    month_data = [month_data; pad];
elseif n_month > n_template
    warning('⚠️ Monthly data has more rows than template; truncating to template height.');
    month_data = month_data(1:n_template, :);
end

vars = template_tbl.Properties.VariableNames;

idx_monthly9 = find(strcmp(vars, 'Monthly_9'), 1);
idx_seasonal = find(strcmp(vars, 'Seasonal'), 1);

if isempty(idx_monthly9)
    error('❌ Cannot find Monthly_9 in Studies and Fluxes-unit.xlsx.');
end
if isempty(idx_seasonal) || idx_seasonal <= idx_monthly9
    % fallback: insert right after Monthly_9
    idx_seasonal = idx_monthly9 + 1;
end

before = template_tbl(:, 1:idx_monthly9);

% middle columns between Monthly_9 and Seasonal (often empty)
if (idx_monthly9+1) <= (idx_seasonal-1)
    middle = template_tbl(:, idx_monthly9+1 : idx_seasonal-1);
else
    middle = template_tbl(:, []); % empty table
end

% after from Seasonal to end
if idx_seasonal <= width(template_tbl)
    after = template_tbl(:, idx_seasonal:end);
else
    after = template_tbl(:, []); % insert at end
end

tbl_updated = [before month_data middle after];
fprintf('✅ Step5: inserted m1..m12 between Monthly_9 and Seasonal.\n');

%% ===========================
% Step 6 (one-sentence summary):
% Extract monthly/seasonal/annual blocks from the updated flux-unit template for unit conversion & aggregation.
%% ===========================

disp('⏳ Step6: extracting blocks (IN MEMORY) ...');

tbl = tbl_updated;

monthly_fields = ["Monthly_4","Monthly_5","Monthly_3","Monthly_6","Monthly_7","Monthly_8","Monthly_9", ...
                  "m1","m2","m3","m4","m5","m6","m7","m8","m9","m10","m11","m12"];

seasonal_fields = ["Seasonal_4","Seasonal_3","Seasonal_5","Seasonal_6","Seasonal_7","Seasonal_8","Seasonal_9", ...
                   "Seasonal_spring","Seasonal_summer","Seasonal_autumn","Seasonal_winter", ...
                   "Seasonal_growing_vegetation","Seasonal_dormant","Seasonal_warm/wet","Seasonal_cool/dry"];

annual_fields = ["Annual_4","Annual_3","Annual_5","Annual_6","Annual_7","Annual_8","Annual_9","Annual_annual"];

extract_vars = @(fields) tbl(:, intersect(fields, tbl.Properties.VariableNames, 'stable'));

monthly  = extract_vars(monthly_fields);
seasonal = extract_vars(seasonal_fields);
annual   = extract_vars(annual_fields);

%% ===========================
% Step 7 (one-sentence summary):
% Convert monthly/seasonal/annual flux values into a unified unit using metadata-defined conversion factors.
%% ===========================

disp('⏳ Step7: unit conversion (IN MEMORY) ...');

normalize_unit = @(s) lower(strtrim(regexprep(string(s), '[µμ]', 'u')));

mass_units = containers.Map( ...
    {'ug','ng','mg','g','kg','t','nmol','umol','mmol','mol','molecules'}, ...
    [1,   1e-3,1e3,1e6,1e9,1e12,16e-3, 16,   16e3, 16e6, 2.66e-17]);

area_units = containers.Map({'m2','ha','hm2','cm2'}, [1,1e4,1e4,1e-4]);

time_units = containers.Map({'h','hr','hour','min','s','sec','d','day','mon','month','y','yr','year','season','sea'}, ...
    [1, 1,    1,     1/60, 1/3600, 1/3600, 24, 24, 24*30, 24*30, 24*365, 24*365, 24*365, 24*90, 24*90]);

get_gas_factor = @(gas) ...
    (strcmpi(gas, 'co2') || strcmpi(gas, 'co2-eq')) * (1/27.2 * 0.75) + ...
    strcmpi(gas, 'ch4') * 0.75 + ...
    (~(strcmpi(gas, 'ch4') || strcmpi(gas, 'co2') || strcmpi(gas, 'co2-eq'))) * 1.0;

compute_unit_factor = @(gas, mass_u, scale3, area_u, time_u, time_multi, signv) ...
    get_val(mass_units, mass_u) .* ...
    eval_safe(scale3) ./ ...
    get_val(area_units, area_u) ./ ...
    (get_val(time_units, time_u) .* eval_safe(time_multi)) .* ...
    get_gas_factor(gas) .* ...
    eval_safe(signv);

% ---- Monthly conversion ----
mu    = normalize_unit(monthly.Monthly_5);
gas   = lower(string(monthly.Monthly_4));
scale = monthly.Monthly_3;
au    = normalize_unit(monthly.Monthly_6);
tu    = normalize_unit(monthly.Monthly_7);
tm    = monthly.Monthly_8;
signv = monthly.Monthly_9;

conv_factor = arrayfun(@(i) compute_unit_factor(gas(i), mu(i), scale(i), au(i), tu(i), tm(i), signv(i)), ...
    (1:height(monthly))');

monthly.Monthly_conv_factor = conv_factor;

for m = 1:12
    col = sprintf('m%d', m);
    conv_col = sprintf('m%d_conv', m);
    if ismember(col, monthly.Properties.VariableNames)
        monthly.(conv_col) = f4_numify_col(monthly.(col)) .* conv_factor;
    end
end

% ---- Seasonal conversion ----
mu    = normalize_unit(seasonal.Seasonal_5);
gas   = lower(string(seasonal.Seasonal_4));
scale = seasonal.Seasonal_3;
au    = normalize_unit(seasonal.Seasonal_6);
tu    = normalize_unit(seasonal.Seasonal_7);
tm    = seasonal.Seasonal_8;
signv = seasonal.Seasonal_9;

conv_factor_s = arrayfun(@(i) compute_unit_factor(gas(i), mu(i), scale(i), au(i), tu(i), tm(i), signv(i)), ...
    (1:height(seasonal))');

seasonal.Seasonal_conv_factor = conv_factor_s;

colsS = ["Seasonal_spring","Seasonal_summer","Seasonal_autumn","Seasonal_winter", ...
         "Seasonal_growing_vegetation","Seasonal_dormant","Seasonal_warm/wet","Seasonal_cool/dry"];

for i = 1:numel(colsS)
    col = colsS(i);
    new_col = col + "_conv";
    if ismember(col, seasonal.Properties.VariableNames)
        seasonal.(new_col) = f4_numify_col(seasonal.(col)) .* conv_factor_s;
    end
end

% ---- Annual conversion ----
mu    = normalize_unit(annual.Annual_5);
gas   = lower(string(annual.Annual_4));
scale = annual.Annual_3;
au    = normalize_unit(annual.Annual_6);
tu    = normalize_unit(annual.Annual_7);
tm    = annual.Annual_8;
signv = annual.Annual_9;

conv_factor_a = arrayfun(@(i) compute_unit_factor(gas(i), mu(i), scale(i), au(i), tu(i), tm(i), signv(i)), ...
    (1:height(annual))');

annual.Annual_conv_factor = conv_factor_a;

if ismember("Annual_annual", annual.Properties.VariableNames)
    annual.Annual_annual_conv = f4_numify_col(annual.Annual_annual) .* conv_factor_a;
end

%% ===========================
% Step 8 (one-sentence summary):
% Compute seasonal and annual means from converted monthly/seasonal values using strict availability rules.
%% ===========================

disp('⏳ Step8: calculating seasonal and annual means (IN MEMORY) ...');

monthly.M_Spring = NaN(height(monthly),1);
monthly.M_Summer = NaN(height(monthly),1);
monthly.M_Autumn = NaN(height(monthly),1);
monthly.M_Winter = NaN(height(monthly),1);
monthly.M_Annual = NaN(height(monthly),1);

for i = 1:height(monthly)
    mvals = NaN(1,12);
    for m = 1:12
        col = sprintf('m%d_conv', m);
        if ismember(col, monthly.Properties.VariableNames)
            mvals(m) = monthly.(col)(i);
        end
    end

    if all(~isnan(mvals([3 4 5]))),   monthly.M_Spring(i) = mean(mvals([3 4 5])); end
    if all(~isnan(mvals([6 7 8]))),   monthly.M_Summer(i) = mean(mvals([6 7 8])); end
    if all(~isnan(mvals([9 10 11]))), monthly.M_Autumn(i) = mean(mvals([9 10 11])); end
    if all(~isnan(mvals([12 1 2]))),  monthly.M_Winter(i) = mean(mvals([12 1 2])); end

    if all(~isnan(mvals))
        monthly.M_Annual(i) = mean(mvals);
    end
end

% Seasonal annual mean priority rules
seasonal.Sea_Spring_conv = seasonal.Seasonal_spring_conv;
seasonal.Sea_Summer_conv = seasonal.Seasonal_summer_conv;
seasonal.Sea_Autumn_conv = seasonal.Seasonal_autumn_conv;
seasonal.Sea_Winter_conv = seasonal.Seasonal_winter_conv;
seasonal.Sea_Annual = NaN(height(seasonal),1);

for i = 1:height(seasonal)
    A = seasonal.Seasonal_spring_conv(i);
    B = seasonal.Seasonal_summer_conv(i);
    C = seasonal.Seasonal_autumn_conv(i);
    D = seasonal.Seasonal_winter_conv(i);

    if all(~isnan([A,B,C,D]))
        seasonal.Sea_Annual(i) = mean([A,B,C,D]);
        continue;
    end

    G = seasonal.Seasonal_growing_vegetation_conv(i);
    H = seasonal.Seasonal_dormant_conv(i);
    if all(~isnan([G,H]))
        seasonal.Sea_Annual(i) = mean([G,H]);
        continue;
    end

    if ismember("Seasonal_warm/wet_conv", seasonal.Properties.VariableNames) && ...
       ismember("Seasonal_cool/dry_conv", seasonal.Properties.VariableNames)
        U = seasonal.("Seasonal_warm/wet_conv")(i);
        X = seasonal.("Seasonal_cool/dry_conv")(i);
        if all(~isnan([U,X]))
            seasonal.Sea_Annual(i) = mean([U,X]);
            continue;
        end
    end
end

%% ===========================
% Step 9 (one-sentence summary):
% Build the standardized Final17 table (Jan–Dec + 4 seasons + annual) using annual > seasonal > monthly priority rules.
%% ===========================

disp('📊 Step9: generating Final17 output ...');

n = height(monthly);
Final17 = array2table(NaN(n,17), ...
    'VariableNames', {'Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec', ...
                      'Spring','Summer','Autumn','Winter','Annual'});

for m = 1:12
    mon_str = datestr(datetime(1,m,1), 'mmm'); % Jan..Dec
    col = sprintf('m%d_conv', m);
    if ismember(col, monthly.Properties.VariableNames)
        Final17.(mon_str) = monthly.(col);
    end
end

season_names         = {'Spring','Summer','Autumn','Winter'};
monthly_season_vars  = {'M_Spring','M_Summer','M_Autumn','M_Winter'};
seasonal_season_vars = {'Sea_Spring_conv','Sea_Summer_conv','Sea_Autumn_conv','Sea_Winter_conv'};

for i = 1:4
    if ismember(seasonal_season_vars{i}, seasonal.Properties.VariableNames)
        Final17.(season_names{i}) = seasonal.(seasonal_season_vars{i});
    end
    idx_nan = isnan(Final17.(season_names{i})) & ismember(monthly_season_vars{i}, monthly.Properties.VariableNames);
    Final17.(season_names{i})(idx_nan) = monthly.(monthly_season_vars{i})(idx_nan);
end

Final17.Annual = NaN(n,1);
if ismember("Annual_annual_conv", annual.Properties.VariableNames)
    Final17.Annual = annual.Annual_annual_conv;
end
idx_nan = isnan(Final17.Annual) & ismember("Sea_Annual", seasonal.Properties.VariableNames);
Final17.Annual(idx_nan) = seasonal.Sea_Annual(idx_nan);
idx_nan = isnan(Final17.Annual) & ismember("M_Annual", monthly.Properties.VariableNames);
Final17.Annual(idx_nan) = monthly.M_Annual(idx_nan);

disp('✅ Step9: Final17 generated in memory (no Excel output written).');

%% ===========================
% Step 10 (one-sentence summary):
% Paste Final17 back into Studies and Fluxes.xlsx by (Paper_number, Study_number) matching
% BUT ONLY write the 17 target columns back (NO full-table overwrite, preserve Q columns).
%% ===========================

disp('📌 Step10: pasting Final17 into Studies and Fluxes.xlsx (PARTIAL WRITE ONLY 17 COLS, KEEP ROW ORDER) ...');

needTargetCols = ["Paper_number","Study_number","Jan","Dec","Annual"];
[target_sheet, T_all] = read_sheet_with_cols(TARGET_FLUX_XLSX, needTargetCols);
fprintf('✅ Step10: target sheet used: %s\n', target_sheet);

targetCols17 = {'Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec', ...
                'Spring','Summer','Autumn','Winter','Annual'};

missAll = targetCols17(~ismember(targetCols17, T_all.Properties.VariableNames));
assert(isempty(missAll), '❌ Target file missing these columns (must exist): %s', strjoin(missAll, ', '));

assert(all(ismember({'Paper_number','Study_number'}, T_all.Properties.VariableNames)), ...
    '❌ Target file must contain Paper_number and Study_number.');

% ---- source keys: use template_tbl numbering (same length as inserted & processed) ----
nUse = min(height(Final17), height(template_tbl));
K_src = template_tbl(1:nUse, {'Paper_number','Study_number'});
F_use = Final17(1:nUse, targetCols17);

K_src.Paper_number = toStringKey_step10(K_src.Paper_number);
K_src.Study_number = toStringKey_step10(K_src.Study_number);

T_all_key_p = toStringKey_step10(T_all.Paper_number);
T_all_key_s = toStringKey_step10(T_all.Study_number);

srcKeyAll = K_src.Paper_number + "|" + K_src.Study_number;

% deduplicate src keys: keep last occurrence
[~, ia_last] = unique(flipud(srcKeyAll), 'stable');
ia_last = nUse - ia_last + 1;
ia_last = sort(ia_last);

K_src = K_src(ia_last, :);
F_use = F_use(ia_last, :);
srcKey = K_src.Paper_number + "|" + K_src.Study_number;

allKey = T_all_key_p + "|" + T_all_key_s;
[found, idxInSrc] = ismember(allKey, srcKey);

nOverwrite  = 0;
nRowMatched = nnz(found);

% ---- update values in memory (T_all) ----
for ii = 1:numel(targetCols17)
    c = targetCols17{ii};

    oldData = T_all.(c);
    srcCol  = F_use.(c);

    vals = NaN(height(T_all),1);
    idxVec = idxInSrc(found);

    okIdx = idxVec > 0 & idxVec <= height(F_use);
    tmpFound = false(height(T_all),1);
    tmpFound(found) = okIdx;

    vals(tmpFound) = srcCol(idxVec(okIdx));

    m = tmpFound & ~isnan(vals);
    nOverwrite = nOverwrite + nnz(m);

    oldData(m) = vals(m);
    T_all.(c) = oldData;
end

% ✅ 可选备份（默认关闭）
if DO_BACKUP
    bk = fullfile(ROOT, ['Studies_and_Fluxes_backup_' datestr(now,'yyyymmdd_HHMMSS') '.xlsx']);
    try
        copyfile(TARGET_FLUX_XLSX, bk);
        fprintf('✅ Backup created: %s\n', bk);
    catch
        warning('⚠️ Backup failed (copyfile). Continue writing anyway...');
    end
end

% ✅ 关键修复：不要 writetable(T_all,...)（会重写整表，导致 √ 列变化）
% ✅ 只把 17 列写回原来的位置

colStart = find(strcmp(T_all.Properties.VariableNames, 'Jan'), 1, 'first');
assert(~isempty(colStart), '❌ Cannot locate Jan column index in target sheet.');
startCell = sprintf('%s2', colnum2excel(colStart));  % row 1 header, data start row 2

% build block
try
    blk = table2array(T_all(:, targetCols17));
catch
    blk = cellfun(@f4_numify, table2cell(T_all(:, targetCols17)));
end

% write as cell (NaN -> blank), better to avoid Excel showing "NaN"
Cblk = num2cell(blk);
Cblk(isnan(blk)) = {[]};

writecell(Cblk, TARGET_FLUX_XLSX, 'Sheet', target_sheet, 'Range', startCell);

fprintf('✅ Step10 done (PARTIAL WRITE ONLY 17 COLS).\n');
fprintf('   Rows matched: %d\n', nRowMatched);
fprintf('   Cells overwritten (non-NaN only): %d\n', nOverwrite);
fprintf('✅ Target updated (Q columns preserved): %s\n', TARGET_FLUX_XLSX);

%% ======================= Function definitions =======================

function y = to_num_safe(x)
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

function s = to_scalar(v)
    if isnumeric(v) || islogical(v)
        if isempty(v), s = NaN; else, s = double(v(1)); end
    else
        s = NaN;
    end
end

function [ok, yy] = parse_year_scalar(x)
    yy = NaN; ok = false;
    if isempty(x)
        return;
    elseif isnumeric(x) && isfinite(x)
        yy = double(x);
        ok = true;
        return;
    elseif ischar(x) || isstring(x)
        ystr = regexp(char(x), '\d{4}', 'match');
        if ~isempty(ystr)
            yy = str2double(ystr{1});
            ok = isfinite(yy);
        end
    end
end

function v = numify_col(v)
    if iscell(v)
        v = cellfun(@numify, v);
    elseif isnumeric(v)
        v = double(v);
    elseif isstring(v) || ischar(v) || iscategorical(v)
        v = str2double(string(v));
    else
        v = arrayfun(@numify, num2cell(v));
    end
end

function y = numify(x)
    if iscell(x)
        if isempty(x), y = NaN;
        elseif isscalar(x), y = numify(x{1});
        else, y = NaN; end
    elseif isnumeric(x)
        if isempty(x), y = NaN;
        elseif isscalar(x), y = double(x);
        else, y = NaN; end
    elseif isstring(x) || ischar(x)
        xs = strtrim(string(x));
        if isscalar(xs), y = str2double(xs); else, y = NaN; end
    elseif isempty(x)
        y = NaN;
    else
        y = NaN;
    end
end

function z = unwrap_scalar_cell(z)
    while iscell(z) && isscalar(z)
        z = z{1};
    end
    if isstring(z) && isscalar(z) && ~ismissing(z)
        z = char(z);
    end
end

function v = f4_numify_col(v)
    if iscell(v)
        v = cellfun(@f4_numify, v);
    elseif isnumeric(v)
        v = double(v);
    elseif isstring(v) || ischar(v) || iscategorical(v)
        v = str2double(string(v));
    else
        v = arrayfun(@f4_numify, num2cell(v));
    end
end

function y = f4_numify(x)
    if iscell(x)
        if isempty(x), y = NaN;
        elseif isscalar(x), y = f4_numify(x{1});
        else, y = NaN; end
    elseif isnumeric(x)
        if isempty(x), y = NaN;
        elseif isscalar(x), y = double(x);
        else, y = NaN; end
    elseif isstring(x) || ischar(x)
        xs = strtrim(string(x));
        if isscalar(xs), y = str2double(xs); else, y = NaN; end
    elseif isempty(x)
        y = NaN;
    else
        y = NaN;
    end
end

function val = get_val(map, key)
    try
        ks = lower(strtrim(string(key)));
        if ismissing(ks) || ks == ""
            val = NaN; return;
        end
        kc = char(ks);
        if map.isKey(kc)
            val = map(kc);
        else
            fprintf('[WARN][UnitMap] Unknown key: %s\n', kc);
            val = NaN;
        end
    catch
        val = NaN;
    end
end

function val = eval_safe(x)
    if isnumeric(x)
        if isempty(x) || ~isfinite(x), val = NaN; else, val = double(x); end
        return;
    end
    xs = strtrim(string(x));
    if xs == "" || ismissing(xs)
        val = NaN; return;
    end
    v = str2double(xs);
    if ~isnan(v)
        val = v; return;
    end
    val = NaN;
end

function s = toStringKey_step10(x)
    if iscell(x)
        s = string(x);
    elseif isnumeric(x)
        s = string(x);
    elseif isstring(x)
        s = x;
    elseif iscategorical(x)
        s = string(x);
    else
        try
            s = string(x);
        catch
            s = string(cellstr(x));
        end
    end
    s = strtrim(s);
end

function [picked_sheet, T] = read_sheet_with_cols(xlsxFile, needCols)
    try
        sh = sheetnames(xlsxFile);
    catch
        [~, sh] = xlsfinfo(xlsxFile);
    end
    if isempty(sh)
        error('❌ No readable sheets: %s', xlsxFile);
    end

    needCols = string(needCols);
    picked_sheet = '';
    T = table();

    for i = 1:numel(sh)
        Ti = readtable(xlsxFile, 'Sheet', sh{i}, 'VariableNamingRule','preserve', 'UseExcel', false);
        vns = string(Ti.Properties.VariableNames);
        if all(ismember(needCols, vns))
            picked_sheet = sh{i};
            T = Ti;
            return;
        end
    end

    error("❌ In file %s, no sheet contains all required columns:\n%s", xlsxFile, strjoin(needCols, ", "));
end

function letters = colnum2excel(n)
    % Convert 1->A, 2->B, ..., 26->Z, 27->AA ...
    letters = '';
    while n > 0
        r = mod(n-1, 26);
        letters = [char(r + 'A') letters];
        n = floor((n-1) / 26);
    end
end
