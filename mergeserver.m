function mergeserver(a)
E = actxserver('Excel.Application');
wb = E.Workbooks;
S = dir('modular-DOT-Table.xlsx');
txt = S.folder;
filedir = strcat(txt, '\modular-DOT-Table.xlsx');
E.DisplayAlerts = 0;
wb = E.Workbooks;
mod = Open(wb, filedir);
sheet = mod.Sheets;
sheet1 = Item(sheet,1);
sheet1 = sheet1.get('Range', a);
sheet1.MergeCells = 1;
mod.SaveAs([cd '\modular-DOT-Table.xlsx']);
Quit(E);
end
