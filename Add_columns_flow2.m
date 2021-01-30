function Add_columns_flow2 (A)
E = actxserver('Excel.Application');
E.DisplayAlerts = 0;
wb = E.Workbooks;
S = dir('modular-DOT-Table.xlsx');
txt = S.folder;
filedir = strcat(txt, '\modular-DOT-Table.xlsx'); 
mod = Open(wb, filedir);
Sheet = mod.Sheets;
Sheet1 = Item(Sheet, 1);
Sheet1 = Sheet1.get('Range', A);
Sheet1.EntireColumn.Insert;
T = A(1);
M = [T '4:' T '5'];
Ss = [T '4'];
X = 'Flow';
Xs = {X};
mod.SaveAs([cd '\modular-DOT-Table.xlsx']);
Quit(E);
mergeserver(M);
end








