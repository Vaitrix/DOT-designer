function Floor_name (X)
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
S = {X};
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B2');
mergeserver('B2:V2');
elseif isfile('modular-DOT-Table.xlsx') == 1
   S = {X};
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B2');
mergeserver('B2:V2'); 
end
