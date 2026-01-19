clear;

%% =========================================================
%  ST ONLY - Scheme A (Recommended)
%  ✅ Only overwrite 12 columns in TARGET Excel (do NOT writetable whole sheet)
%  ✅ No intermediate .mat / .xlsx saved
%  ✅ monout / expanded / filled all in memory
%
%  ROOT:
%   D:\Users\jiaweiChiang\Desktop\Supplymentary_code
%
%  RAW:          CH4 uptake data_CH4 FLUX_1_1410.xlsx  (contains DOY)
%  FLUX-UNIT:    Studies and Fluxes-unit.xlsx         (Paper_number/Study_number ordering)
%  MOVE TEMPLATE move_template_ST.xlsx                (structure only)
%  TARGET:       Studies and Fluxes.xlsx              (final write-back)
% =========================================================

ROOT = 'D:\Users\jiaweiChiang\Desktop\Supplymentary_code';
cd(ROOT);

RAW_XLSX      = fullfile(ROOT, 'CH4 uptake data_CH4 FLUX_1_1410.xlsx');
MOVE_TEMPLATE = fullfile(ROOT, 'move_template_ST.xlsx');

FLUX_UNIT_GUESS = fullfile(ROOT, 'Studies and Fluxes-unit.xlsx');
TARGET_GUESS    = fullfile(ROOT, 'Studies and Fluxes.xlsx');

%% ===================== Step0: checks + locate =====================
fprintf('\n📌 Step0: Input checks + locate files...\n');

assert(isfile(RAW_XLSX),      '❌ Missing RAW_XLSX: %s', RAW_XLSX);
assert(isfile(MOVE_TEMPLATE), '❌ Missing MOVE_TEMPLATE: %s', MOVE_TEMPLATE);

if isfile(FLUX_UNIT_GUESS)
    FLUX_UNIT = FLUX_UNIT_GUESS;
else
    FLUX_UNIT = find_workbook_prefer_pattern_then_cols(ROOT, "*flux-unit*.xlsx", ...
        ["Paper_number","Study_number"]);
end
fprintf('✅ FLUX_UNIT: %s\n', FLUX_UNIT);

if isfile(TARGET_GUESS)
    TARGET_XLSX = TARGET_GUESS;
else
    TARGET_XLSX = find_workbook_prefer_pattern_then_cols_exclude(ROOT, "*.xlsx", "*flux-unit*.xlsx", ...
        ["Paper_number","Study_number"]);
end
fprintf('✅ TARGET_XLSX: %s\n', TARGET_XLSX);

%% =====================================================================
%% [1] DOY -> Monthly means (monout)  —— only in memory
%% =====================================================================
fprintf('\n📌 Step1: DOY -> Monthly means (ST sheet only)...\n');

try
    sheets = sheetnames(RAW_XLSX);
catch
    [~, sheets] = xlsfinfo(RAW_XLSX);
end
if isempty(sheets)
    error('❌ RAW workbook has no readable sheets.');
end

if numel(sheets) < 3
    error('❌ RAW workbook has < 3 sheets (cannot pick ST=sheet3).');
end

skipped_info = {};
monout_cell  = {'NO_DOY'};

% 固定只跑第 3 个 sheet（ST）
s = 3;
sheet = sheets{s};
fprintf('📄 RAW ST sheet = %s\n', sheet);

try
    [~, ~, raw] = xlsread(RAW_XLSX, sheet);
catch
    error('❌ Failed reading RAW sheet: %s', sheet);
end
if isempty(raw)
    error('❌ RAW ST sheet is empty: %s', sheet);
end

[R,C] = size(raw);
output_rows = {};
any_found = false;

for r = 1:(R-1)
    tag = [];
    if C >= 5, tag = raw{r,5}; end

    if ~(ischar(tag) || isstring(tag)) || ~strcmpi(strtrim(char(tag)), 'doy')
        continue;
    end
    any_found = true;

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
    year_calc = round(year_calc);

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
    doy = doy(valid);
    val = val(valid);

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

if any_found && ~isempty(output_rows)
    monout_cell = vertcat(output_rows{:});
    fprintf('✅ monout_cell built in memory: %d rows\n', size(monout_cell,1));
else
    fprintf('⚠️ No DOY markers found → monout_cell=NO_DOY\n');
end

if ~isempty(skipped_info)
    fprintf('⚠️ skipped records: %d (NOT writing any skipped file)\n', size(skipped_info,1));
end

%% =====================================================================
%% [2] Read flux-unit -> build aligned move table (memory only)
%% =====================================================================
fprintf('\n📌 Step2: Build move aligned table (memory only)...\n');

[tpl_sheet, template_tbl] = read_sheet_with_cols(FLUX_UNIT, ["Paper_number","Study_number"]);
fprintf('✅ flux-unit sheet used: %s\n', tpl_sheet);

papers  = str2double(regexprep(string(template_tbl.("Paper_number")),'[^\d\.\-]',''));
studies = str2double(regexprep(string(template_tbl.("Study_number")),'[^\d\.\-]',''));
target_h = numel(papers);

T0 = readtable(MOVE_TEMPLATE, 'VariableNamingRule','preserve','UseExcel',false);
h0 = height(T0);
vnames = T0.Properties.VariableNames;

T2 = table();
for v = 1:numel(vnames)
    vn  = vnames{v};
    col = T0.(vn);

    if isvector(col)
        w = 1; col = col(:);
    else
        w = size(col,2);
    end

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
            tmp = repmat(col(1,:), target_h, 1);
            tmp(:) = categorical(missing);
            T2.(vn) = tmp;
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
    vn = vnames{v};

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

if ~ismember('papers',  string(T2.Properties.VariableNames)), T2.papers  = NaN(target_h,1); end
if ~ismember('studies', string(T2.Properties.VariableNames)), T2.studies = NaN(target_h,1); end
T2.papers  = papers(:);
T2.studies = studies(:);

fprintf('✅ move aligned table built: old=%d rows, new=%d rows\n', h0, height(T2));

%% =====================================================================
%% [3] monout -> compute main-row monthly means -> fill move main rows
%% =====================================================================
fprintf('\n📌 Step3: Compute main-row m1..m12 (avg across layers/years)...\n');

Tmove = T2;
mvNames = Tmove.Properties.VariableNames;

if numel(mvNames) < 16
    error('❌ move template needs >=16 cols: papers,studies,year,layer,m1..m12');
end

mvNames(1:4) = {'papers','studies','year','layer'};
Tmove.Properties.VariableNames = mvNames;
MCOLS = 5:16;

T_main = table();

if iscell(monout_cell) && size(monout_cell,1) >= 2 && ~isequal(monout_cell{1,1},'NO_DOY')
    mon_raw = monout_cell;
    [MR, ~] = size(mon_raw);

    mon_rows = {}; mon_ord = []; ord_ctr = 0;

    for r = 1:2:(MR-1)
        meta = mon_raw(r,:);
        vals = mon_raw(r+1,:);

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
            mvals{k} = vals{1,5+k};
        end

        row16 = cell(1,16);
        row16(1:4)  = {paper,study,year,layer};
        row16(5:16) = mvals;

        ord_ctr = ord_ctr + 1;
        mon_rows(end+1,1:16) = row16; %#ok<AGROW>
        mon_ord(end+1,1)     = ord_ctr; %#ok<AGROW>
    end

    if ~isempty(mon_rows)
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

        T_year = groupsummary(T, {'paper','study','year'}, 'mean', cellstr(colnamesMon(5:end)));
        for m = 1:12
            src = sprintf('mean_m%d',m);
            dst = sprintf('m%d',m);
            T_year.(dst) = T_year.(src);
            T_year.(src) = [];
        end

        varList = arrayfun(@(mm)sprintf('m%d',mm), 1:12, 'UniformOutput', false);
        T_main  = groupsummary(T_year, {'paper','study'}, 'mean', varList);
        for m = 1:12
            src = sprintf('mean_m%d',m);
            dst = sprintf('m%d',m);
            T_main.(dst) = T_main.(src);
            T_main.(src) = [];
        end
    end
end

Tmove.papers  = f4_numify_col(Tmove.papers);
Tmove.studies = f4_numify_col(Tmove.studies);
y_move        = f4_numify_col(Tmove.year);

filled_main_cnt = 0;
if ~isempty(T_main)
    for i = 1:height(Tmove)
        if ~isnan(y_move(i)), continue; end
        pk = Tmove.papers(i);
        sk = Tmove.studies(i);
        if isnan(pk) || isnan(sk), continue; end

        idx = find(T_main.paper == pk & T_main.study == sk, 1, 'first');
        if ~isempty(idx)
            for m = 1:12
                Tmove{i, MCOLS(m)} = T_main{idx, sprintf('m%d',m)};
            end
            filled_main_cnt = filled_main_cnt + 1;
        end
    end
end

fprintf('✅ move main rows filled: %d\n', filled_main_cnt);
final_filled_move = Tmove;

%% =====================================================================
%% [4] SCHEME A: Only write the 12 updated columns into TARGET (NO writetable whole sheet)
%% =====================================================================
fprintf('\n📌 Step4: Scheme A partial-column overwrite to TARGET...\n');

% Read target sheet which contains Paper_number/Study_number
[target_sheet, tgt_tbl] = read_sheet_with_cols(TARGET_XLSX, ["Paper_number","Study_number"]);
fprintf('✅ Target sheet used: %s\n', target_sheet);

vars_tar = string(tgt_tbl.Properties.VariableNames);

% Priority: ST_m1..12 -> m1..12 -> Jan..Dec
useCols = strings(1,12);
modeName = "";

% ST_m*
for m = 1:12
    hit = local_find_col(vars_tar, {sprintf('ST_m%d',m)});
    if ~isempty(hit), useCols(m) = hit; end
end
if all(useCols ~= "")
    modeName = "ST_m1..ST_m12";
else
    % m1..m12
    tmp = strings(1,12);
    for m = 1:12
        hit = local_find_col(vars_tar, {sprintf('m%d',m)});
        if ~isempty(hit), tmp(m) = hit; end
    end
    if all(tmp ~= "")
        useCols = tmp;
        modeName = "m1..m12";
    else
        % Jan..Dec
        janDec = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
        tmp = strings(1,12);
        for m = 1:12
            hit = local_find_col(vars_tar, {char(janDec(m))});
            if ~isempty(hit), tmp(m) = hit; end
        end
        if all(tmp ~= "")
            useCols = tmp;
            modeName = "Jan..Dec";
        else
            error('❌ Target lacks complete ST_m1..12 OR m1..12 OR Jan..Dec (cannot add new columns).');
        end
    end
end
fprintf('✅ Write-back mode: %s\n', modeName);

% Build source map from move main rows
src_tbl = final_filled_move;
src_tbl.papers  = f4_numify_col(src_tbl.papers);
src_tbl.studies = f4_numify_col(src_tbl.studies);
y_src = f4_numify_col(src_tbl.year);

is_main  = isnan(y_src);
src_main = src_tbl(is_main, :);

MCOLS = 5:16;
src_vals = table2cell(src_main(:, MCOLS));
src_vals = cellfun(@f4_numify, src_vals);
src_vals = reshape(src_vals, height(src_main), 12);

mkKey = @(p,s) sprintf('%.15g|%.15g', p, s);
M = containers.Map('KeyType','char','ValueType','any');

for i = 1:height(src_main)
    p = src_main.papers(i);
    s = src_main.studies(i);
    if isnan(p) || isnan(s), continue; end
    M(mkKey(p,s)) = src_vals(i,:);
end

tgt_p = f4_numify_col(tgt_tbl.("Paper_number"));
tgt_s = f4_numify_col(tgt_tbl.("Study_number"));
nRow  = height(tgt_tbl);

UpdatedMat = NaN(nRow, 12);

filled_rows = 0;
for i = 1:nRow
    p = tgt_p(i);
    s = tgt_s(i);
    if isnan(p) || isnan(s), continue; end

    k = mkKey(p,s);
    if isKey(M,k)
        UpdatedMat(i,:) = double(M(k));
        filled_rows = filled_rows + 1;
    end
end

filled_cells = 0;

% ✅ overwrite ONLY 12 cols, column-by-column, start row2 (keep header row)
for m = 1:12
    vn = char(useCols(m));

    oldCol = tgt_tbl.(vn);
    if isnumeric(oldCol)
        newCol = double(oldCol);
    else
        tmp = string(oldCol);
        tmp = strtrim(tmp);
        tmp(tmp=="") = "NaN";
        newCol = str2double(tmp);
    end

    upd = isfinite(UpdatedMat(:,m));
    newCol(upd) = UpdatedMat(upd,m);
    filled_cells = filled_cells + sum(upd);

    colIdx = find(strcmpi(vars_tar, vn), 1, 'first');
    if isempty(colIdx)
        error('❌ cannot locate col index for %s', vn);
    end

    colLetter = local_excel_col(colIdx);
    startCell = sprintf('%s2', colLetter);

    writematrix(newCol, TARGET_XLSX, 'Sheet', target_sheet, 'Range', startCell);
end

fprintf('✅ DONE. Only 12 columns written back.\n');
fprintf('✅ TARGET: %s\n', TARGET_XLSX);
fprintf('➡️ Matched rows: %d\n', filled_rows);
fprintf('➡️ Overwritten cells (finite): %d\n', filled_cells);

%% =====================================================================
%% ========================= Functions =========================
%% =====================================================================

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

function [picked_sheet, T] = read_sheet_with_cols(xlsxFile, needCols)
    try
        sh = sheetnames(xlsxFile);
    catch
        [~, sh] = xlsfinfo(xlsxFile);
    end
    if isempty(sh)
        error('❌ No readable sheets: %s', xlsxFile);
    end

    picked_sheet = '';
    T = table();
    needCols = string(needCols);

    for i = 1:numel(sh)
        Ti = readtable(xlsxFile, 'Sheet', sh{i}, 'VariableNamingRule','preserve', 'UseExcel', false);
        vns = string(Ti.Properties.VariableNames);
        if all(ismember(needCols, vns))
            picked_sheet = sh{i};
            T = Ti;
            return;
        end
    end

    error("❌ In file %s cannot find a sheet containing required columns:\n%s", ...
        xlsxFile, strjoin(needCols, ", "));
end

function xlsxPath = find_workbook_prefer_pattern_then_cols(rootDir, preferPattern, needCols)
    hit = dir(fullfile(rootDir, preferPattern));
    if ~isempty(hit)
        xlsxPath = fullfile(hit(1).folder, hit(1).name);
        return;
    end
    xlsxPath = find_workbook_with_cols(rootDir, needCols);
end

function xlsxPath = find_workbook_prefer_pattern_then_cols_exclude(rootDir, preferPattern, excludePattern, needCols)
    hits = dir(fullfile(rootDir, preferPattern));
    if ~isempty(hits)
        for i = 1:numel(hits)
            f = fullfile(hits(i).folder, hits(i).name);
            if ~isempty(dir(fullfile(rootDir, excludePattern))) && contains(lower(f), lower(strrep(excludePattern,'*','')))
                continue;
            end
            xlsxPath = f;
            return;
        end
    end

    files = dir(fullfile(rootDir, '*.xlsx'));
    needCols = string(needCols);
    for i = 1:numel(files)
        f = fullfile(files(i).folder, files(i).name);
        if ~isempty(dir(fullfile(rootDir, excludePattern))) && contains(lower(f), lower(strrep(excludePattern,'*','')))
            continue;
        end
        try
            try
                sh = sheetnames(f);
            catch
                [~, sh] = xlsfinfo(f);
            end
            if isempty(sh), continue; end
            for s = 1:numel(sh)
                Ti = readtable(f, 'Sheet', sh{s}, 'VariableNamingRule','preserve', 'UseExcel', false);
                vns = string(Ti.Properties.VariableNames);
                if all(ismember(needCols, vns))
                    xlsxPath = f;
                    return;
                end
            end
        catch
            continue;
        end
    end

    error("❌ Cannot locate workbook under %s (prefer=%s, cols=%s)", ...
        rootDir, preferPattern, strjoin(needCols, ", "));
end

function xlsxPath = find_workbook_with_cols(rootDir, needCols)
    files = dir(fullfile(rootDir, '*.xlsx'));
    if isempty(files)
        error('❌ No xlsx files in %s', rootDir);
    end
    needCols = string(needCols);

    for i = 1:numel(files)
        f = fullfile(files(i).folder, files(i).name);
        try
            try
                sh = sheetnames(f);
            catch
                [~, sh] = xlsfinfo(f);
            end
            if isempty(sh), continue; end

            for s = 1:numel(sh)
                Ti = readtable(f, 'Sheet', sh{s}, 'VariableNamingRule','preserve', 'UseExcel', false);
                vns = string(Ti.Properties.VariableNames);
                if all(ismember(needCols, vns))
                    xlsxPath = f;
                    return;
                end
            end
        catch
            continue;
        end
    end

    error("❌ Cannot find workbook containing columns:\n%s", strjoin(needCols, ", "));
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

function col = local_excel_col(n)
    col = "";
    while n > 0
        r = mod(n-1, 26);
        col = char('A' + r) + col;
        n = floor((n-1)/26);
    end
    col = char(col);
end