function FLOOR (y)
C = {y};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates an excel spreadsheet to load data on from a template.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data to Excel spreadsheet
writecell(C, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B6');fl
elseif isfile('modular-DOT-Table.xlsx') == 1 && isfile('~$modular-DOT-Table.xlsx') == 0
    