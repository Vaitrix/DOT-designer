function condense
persistent String
String = cell(1,10000);
X = zeros(1,10000);
n = 0;
x = DOT_data.index;
for l = 6:x.index_prime
    N = num2str(l);
    Str = ['B' N ];
    Str1 = ['B' N ', '];
   [~,~,K] = xlsread('modular-DOT-Table.xlsx',1,Str);
    if isnan(K{1,1})
        n = n+1;
       String{1,n} = Str1;
       X(1,n) = 1;
    end
 Y = logical(X); 
 D = String;
 d = D(Y);
 M = cell2mat(d);
 [~,w] = size(M);
 Cs = M(1, 1:w-2);
end
E = actxserver('Excel.Application');
E.DisplayAlerts = 0;
E.Visible = 0;
work = E.Workbooks; 
S = dir('modular-DOT-Table.xlsx');
txt = S.folder;
filedir = strcat(txt, '\modular-DOT-Table.xlsx');
wb = Open(work, filedir);
sheet = wb.Sheets;
sheet1 = Item(sheet, 1);
sheet1 = sheet1.Range(Cs);
sheet1.EntireRow.Delete;
wb.SaveAs([cd '\modular-DOT-Table.xlsx']);
Quit(E);
end

    

