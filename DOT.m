classdef DOT
    %DOT_log all the functions required for the DOT Table.
     properties      
     end   
 methods (Static)
function anaesthetic_room (Q,B,X,I)
%Need to make ldr a re-executable function.
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{Q*B},{[Q*B]},{[]},{Q*B},{[]},{Q*B}, {Q*B}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:K7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'N7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow('Q6'); 
Eq = 130+[(B-1)*130/8];
AGS(I,B,Eq);
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{Q*B},{[Q*B]},{[]},{Q*B},{[]},{Q*B},{Q*B}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':K' np];
C = ['C' np];
D = ['D' np];
N = ['N' np];
BB = ['B' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', N);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
Eq = 130+[(B-1)*130/8];
AGS(I,B,Eq);
else
    error('FILE ERROR: potential file overwrite or open file');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function Anaesthetist_aadca(Q,B,X,I)
%Need to make ldr a re-executable function.
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{Q*B},{[Q*B]},{[]},{[Q*B]},{[]},{[Q*B]}, {[Q*B]}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:K7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'N7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow('Q6');
Eq = 130+[(B-1)*130/4];
AGS(I,B,Eq);
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{Q*B},{[Q*B]},{[]},{[Q*B]},{[]},{Q*B},{Q*B}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':K' np];
C = ['C' np];
D = ['D' np];
N = ['N' np];
BB = ['B' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', N);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
Eq = 130+[(B-1)*130/4];
AGS(I,B,Eq);
else
    error('FILE ERROR: potential file overwrite');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function angiography (Q,B,X,I)
%Need to make ldr a re-executable function.
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{B},{[B]},{[]},{[B]},{[]},{[B]}, {[B]}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:K7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'N7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;         
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10 + [(B-1)*(6/3)];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Nitrous Oxide
q = 10 + [(B-1)*(6/4)];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'G7');
%For Medical Air
q = 40 + [(B-1)*40/4];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7')
%For Vacuum
q = 40 + [(B-1)*40/8];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7')
Eq = 130+[(B-1)*130/4];
AGS(I,B,Eq);
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{B},{[]},{[B]},{[]},{[]},{[]},{[B]},{[]},{[]},{[]},{B},{[]},{B}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
UU = ['U' np];
BB = ['B' np];
E = ['E' np];
G = ['G' np];
K = ['K' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', UU);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);         
%For Oxygen
q = 10 + [(B-1)*(6/3)];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Nitrous Oxide
q = 10 + [(B-1)*(6/4)];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', G);
%For Medical Air
q = 40 + [(B-1)*40/4];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K);
%For Vacuum
q = 40 + [(B-1)*40/8];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
Eq = 130+[(B-1)*130/4];
AGS(I,B,Eq);
else
    error('FILE ERROR: potential file overwrite');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end 
 function burns_unit(Q,B,X,I)
%Need to make ldr a re-executable function.
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{2*B},{[2*B]},{[2*B]},{2*B},{[]},{2*B}, {2*B}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:K7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
writematrix(2, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'N7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10+[(B-1)*3*(6/4)];
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7')
%For Nitrogen
q = 10+(B-1)*(6/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'G7')
%For Entonox
q = 20+(B-1)*(10/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'I7')
%For Medical Air
q = 80+(B-1)*(80/2);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7')
%For Vacuum
q = 40+(B-1)*(40/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7')
Eq = 130+[(B-1)*130/8];
AGS(I,B,Eq);
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{2*B},{[]},{2*B},{[]},{2*B},{[]},{2*B},{[]},{[]},{[]},{2*B},{[]},{2*B}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
U = ['U' np];
BB = ['B' np];
E = ['E' np];
G = ['G' np];
I = ['I' np];
K = ['K' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(2, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', U);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
%For Oxygen
q = 10+[(B-1)*3*(6/4)];
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Nitrogen
q = 10+(B-1)*(6/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', G);
%For Entonox
q = 20+(B-1)*(10/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', I);
%For Medical Air
q = 80+(B-1)*(80/2);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K);
%For Vacuum
q = 40+(B-1)*(40/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
Eq = 130+[(B-1)*130/8];
AGS(I,B,Eq);
else
    error('FILE ERROR: potential file overwrite or open file');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
 function CAT(Q,B,X,I)
%Need to make ldr a re-executable function.
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{B},{[B]},{[]},{[B]},{[]},{[B]}, {[B]}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:K7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'M7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10 + [(B-1)*(6/3)];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Nitrous Oxide
q = 10 + [(B-1)*(6/4)];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'G7');
%For Medical Air
q = 40 + [(B-1)*40/4];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7')
%For Vacuum
q = 40 + [(B-1)*40/8];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7')
Eq = 130+[(B-1)*130/4];
AGS(I,B,Eq);
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{B},{[]},{[B]},{[]},{[]},{[]},{B},{[]},{[]},{[]},{B},{[]},{B}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
TT = ['T' np];
BB = ['B' np];
E = ['E' np];
G = ['G' np];
K = ['K' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', TT);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
%For Oxygen
q = 10 + [(B-1)*(6/3)];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Nitrous Oxide
q = 10 + [(B-1)*(6/4)];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', G);
%For Medical Air
q = 40 + [(B-1)*40/4];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K);
%For Vacuum
q = 40 + [(B-1)*40/8];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
Eq = 130+[(B-1)*130/4];
AGS(I,B,Eq);
else
    error('FILE ERROR: potential file overwrite');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
 end
function ccu(Q,B,X,I)
%Need to make ldr a re-executable function.
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{4*B},{[]},{[]},{4*B},{[]},{4*B},{4*B}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:K7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(2, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'L7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10 + [(B-1)*(6*(3/4))];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Medical Air
q = 80 + [(B-1)*80/2];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7');
%For Vacuum
q = 40 + [(B-1)*40/4];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
Eq = 130+[(B-1)*130/8];
AGS(I,B,Eq);
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{4*B},{[]},{[]},{[]},{[]},{[]},{4*B},{[]},{[]},{[]},{4*B},{[]},{4*B}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
SS = ['S' np];
BB = ['B' np];
E = ['E' np];
K = ['K' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(2, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', SS);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
%For Oxygen
q = 10 + [(B-1)*(6*(3/4))];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Medical Air
q = 80 + [(B-1)*80/2];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K);
%For Vacuum
q = 40 + [(B-1)*40/4];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
Eq = 130+[(B-1)*130/4];
AGS(I,B,Eq);
else
    error('FILE ERROR: potential file overwrite or open file');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function critical_care_area(Q,B,X,I)
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{4*B},{[]},{[]},{4*B},{[]},{4*B},{[]}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:K7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(2, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'M7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10 + [(B-1)*(6*(3/4))];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Medical Air
q = 80 + [(B-1)*80/2];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7');
%For Vacuum
q = 40 + [(B-1)*40/4];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
Eq = 130+[(B-1)*130/8];
AGS(I,B,Eq);
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{4*B},{[]},{[]},{[]},{[]},{[]},{4*B},{[]},{[]},{[]},{4*B},{[]}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
TT = ['T' np];
BB = ['B' np];
E = ['E' np];
K = ['K' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(2, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', TT);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
%For Oxygen
q = 10 + [(B-1)*(6*(3/4))];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Medical Air
q = 80 + [(B-1)*80/2];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K);
%For Vacuum
q = 40 + [(B-1)*40/4];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
Eq = 130+[(B-1)*130/8];
AGS(I,B,Eq);
else
    error('FILE ERROR: potential file overwrite or open file');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function day_care(Q,B,X,I)
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{B},{[]},{[]},{[]},{[]},{B},{[]}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:K7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'L7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10 + ((B-1)*(6/4));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Vacuum
q = 40 + (B-1)*40/4;
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
Add_columns_flow('Q6');
Eq = 130+[(B-1)*130/8];
AGS(I,B,Eq);
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{B},{[]},{[]},{[]},{[]},{[]},{[]},{[]},{[]},{[]},{B},{[]}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
SS = ['S' np];
BB = ['B' np];
E = ['E' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', SS);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
%For Oxygen
q = 10 + (B-1)*(6*(3/4));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Vacuum
q = 40 + (B-1)*40/4;
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
Eq = 130+[(B-1)*130/4];
AGS(I,B,Eq);
else
    error('FILE ERROR: potential file overwrite or open file');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function ECT (Q,B,X,I)
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{B},{B},{[]},{B},{[]},{B}, {B}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:K7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'N7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10 + (B-1)*(6/4);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Nitrous Oxide
q = 10;
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'G7');
%For Medical Air
q = 40;
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7');
%For Vacuum
q = 40 + [(B-1)*40/4];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
Eq = 130+[(B-1)*130/8];
AGS(I,B,Eq);
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{B},{[]},{[B]},{[]},{[]},{[]},{B},{[]},{[]},{[]},{B},{[]},{B}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
U = ['U' np];
BB = ['B' np];
E = ['E' np];
G = ['G' np];
K = ['K' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', U);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
%For Oxygen
q = 10 + (B-1)*(6/4);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Nitrous Oxide
q = 10;
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', G);
%For Medical Air
q = 40;
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K);
%For Vacuum
q = 40 + [(B-1)*40/4];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
Eq = 130+[(B-1)*130/8];
AGS(I,B,Eq);
else
    error('FILE ERROR: potential file overwrite or open file');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function endoscopy(Q,B,X,I)
%Need to make ldr a re-executable function.
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{B},{[B]},{[]},{[B]},{[]},{[B]}, {[B]}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:K7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'N7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10 + (B-1)*(6/4);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Nitrous Oxide
q = 10+(B-1)*(6/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'G7');
%For Medical Air
q = 40+(B-1)*(40/4);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7');
%For Vacuum
q = 40 + [(B-1)*40/4];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
Eq = 130+[(B-1)*130/4];
AGS(I,B,Eq);
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{B},{[]},{[B]},{[]},{[]},{[]},{B},{[]},{[]},{[]},{B},{[]},{B}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
U = ['U' np];
BB = ['B' np];
E = ['E' np];
G = ['G' np];
K = ['K' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', U);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
%For Oxygen
q = 10 + (B-1)*(6/4);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Nitrous Oxide
q = 10+(B-1)*(6/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', G);
%For Medical Air
q = 40+(B-1)*(40/4);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K);
%For Vacuum
q = 40 + [(B-1)*40/8];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
Eq = 130+[(B-1)*130/4];
AGS(I,B,Eq);
else
    error('FILE ERROR: potential file overwrite');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function Equipment_room(Q,B,X,I)
%Need to make ldr a re-executable function.
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{B},{B},{[]},{B},{B},{B}, {B}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:K7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 100;
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Nitrous Oxide
q = 15;
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'G7');
%For Medical Air
q = 40;
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7');
%For Surgical Air
q = 350;
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','M7')
%For Vacuum
q = 40;
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
Eq = 130+[(B-1)*130/8];
AGS(I,B,Eq);
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{B},{[]},{B},{[]},{[]},{[]},{B},{[]},{B},{[]},{B},{[]},{B}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
M = ['M' np];
V = ['V' np];
BB = ['B' np];
E = ['E' np];
G = ['G' np];
K = ['K' np];
M = ['M' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', V);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
%For Oxygen
q = 100;
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Nitrous Oxide
q = 15;
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', G);
%For Medical Air
q = 40;
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K);
%For Surgical Air
q = 350;
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', M)
%For Vacuum
q = 40;
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
Eq = 130+[(B-1)*130/8];
AGS(I,B,Eq);
else
    error('FILE ERROR: potential file overwrite');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function Equipment_room_NICU(Q,B,X,I)
%Need to make ldr a re-executable function.
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{B},{[]},{[]},{B},{[]},{B}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:J7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'M7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10;
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Medical Air
q = 40;
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7');
%For Vacuum
q = 40;
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
Eq = 130+[(B-1)*130/8];
AGS(I,B,Eq);
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{B},{[]},{[]},{[]},{[]},{[]},{B},{[]},{[]},{[]},{B},{[]}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
TT = ['TT' np];
BB = ['B' np];
E = ['E' np];
K = ['K' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', TT);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
%For Oxygen
q = 10;
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Medical Air
q = 40;
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K);
%For Vacuum
q = 40;
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
Eq = 130+[(B-1)*130/8];
AGS(I,B,Eq);
else
    error('FILE ERROR: potential file overwrite or open file');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function fluoroscopy (Q,B,X,I)
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{B},{[]},{[]},{[]},{[]},{B}, {[]}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:J7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'L7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10+((B-1)*(6/3));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Vacuum
q = 40+((B-1)*40/8);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
Eq = 130+[(B-1)*130/4];
AGS(I,B,Eq);
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{B},{[]},{[]},{[]},{[]},{[]},{[]},{[]},{[]},{[]},{B},{[]}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
SS = ['S' np];
BB = ['B' np];
E = ['E' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', SS);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
%For Oxygen
q = 10+((B-1)*(6/3));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Vacuum
q = 40 +((B-1)*(40/8));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
Eq = 130+[(B-1)*130/4];
AGS(I,B,Eq);
else
    error('FILE ERROR: potential file overwrite or open file');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function general_purpose (Q,B,X,I)
 % Used flow calculation for inpatient rooms.
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{B},{[]},{[]},{[]},{[]},{B}, {B}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:J7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'L7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10+((B-1)*(6/4));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Vacuum
q = 40+((B-1)*40/4);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
Eq = 130+[(B-1)*130/8];
AGS(I,B,Eq);
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{B},{[]},{[]},{[]},{[]},{[]},{[]},{[]},{[]},{[]},{B},{[]}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
SS = ['S' np];
BB = ['B' np];
E = ['E' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', SS);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
%For Oxygen
q = 10+((B-1)*(6/4));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Vacuum
q = 40 +((B-1)*(40/4));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
Eq = 130+[(B-1)*130/8];
AGS(I,B,Eq);
else
    error('FILE ERROR: potential file overwrite');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function HDU(Q,B,X,I)    
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{4*B},{[]},{[]},{[4*B]},{[]},{4*B}, {[]}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:J7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'M7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10+((B-1)*3*(6/4));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Medical Air
q = 80+((B-1)*80/2);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7');
%For Vacuum
q = 40+((B-1)*40/4);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
Eq = 130+[(B-1)*130/8];
AGS(I,B,Eq);
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{4*B},{[]},{[]},{[]},{[]},{[]},{4*B},{[]},{[]},{[]},{4*B},{[]}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
TT = ['T' np];
BB = ['B' np];
E = ['E' np];
K = ['K' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', TT);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
%For Oxygen
q = 10+((B-1)*3*(6/4));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Medical Air
q = 80+((B-1)*80/2);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K);
%For Vacuum
q = 40+((B-1)*40/4);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
Eq = 130+[(B-1)*130/8];
AGS(I,B,Eq);
else
    error('FILE ERROR: potential file overwrite');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function Holding_and_Recovery(Q,B,X,I)
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{4*B},{[]},{[]},{4*B},{[]},{4*B}, {[]}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:J7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'M7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10+((B-1)*3*(6/4));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Medical Air
q = 80+((B-1)*80/2);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7');
%For Vacuum
q = 40+((B-1)*40/4);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
Eq = 130+[(B-1)*130/8];
AGS(I,B,Eq);
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{4*B},{[]},{[]},{[]},{[]},{[]},{4*B},{[]},{[]},{[]},{4*B},{[]}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
TT = ['T' np];
BB = ['B' np];
E = ['E' np];
K = ['K' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', TT);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
%For Oxygen
q = 10+((B-1)*3*(6/4));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Medical Air
q = 80+((B-1)*80/2);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K);
%For Vacuum
q = 40+((B-1)*40/4);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
Eq = 130+[(B-1)*130/8];
AGS(I,B,Eq);
else
    error('FILE ERROR: potential file overwrite');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function IAP_room(Q,B,X)
%Need to make ldr a re-executable function.
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{[Q*B]},{[]},{[]},{Q*B},{[Q*B]},{[]}, {[]}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:K7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'M7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[B]},{[]},{[]},{B},{[B]},{[]}, {[]}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':K' np];
C = ['C' np];
D = ['D' np];
M = ['M' np];
BB = ['B' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', M);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
else
    error('FILE ERROR: potential file overwrite or open file');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function icu(Q,B,X,I)
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{2*B},{[]},{[]},{[2*B]},{[]},{2*B}, {[]}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:J7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(2, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'M7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10+((B-1)*3*(6/4));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Medical Air
q = 80+((B-1)*80/2);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7');
%For Vacuum
q = 40+((B-1)*40/4);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
Eq = 130+[(B-1)*130/8];
AGS(I,B,Eq);
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{2*B},{[]},{[]},{[]},{[]},{[]},{2*B},{[]},{[]},{[]},{2*B},{[]}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
TT = ['T' np];
BB = ['B' np];
E = ['E' np];
K = ['K' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(2, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', TT);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
%For Oxygen
q = 10+((B-1)*3*(6/4));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Medical Air
q = 80+((B-1)*80/2);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K);
%For Vacuum
q = 40+((B-1)*40/4);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
Eq = 130+[(B-1)*130/8];
AGS(I,B,Eq);
else
    error('FILE ERROR: potential file overwrite');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function Inpatient_room(Q,B,X)
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{B},{[]},{[]},{[B]},{[]},{B}, {[]}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:J7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'M7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10+((B-1)*(6/4));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Medical Air
q = 20+((B-1)*10/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7');
%For Vacuum
q = 40+((B-1)*40/4);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{B},{[]},{[]},{[]},{[]},{[]},{B},{[]},{[]},{[]},{B},{[]}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
TT = ['T' np];
BB = ['B' np];
E = ['E' np];
K = ['K' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', TT);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
%For Oxygen
q = 10+((B-1)*(6/4));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Medical Air
q = 20+((B-1)*10/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K);
%For Vacuum
q = 40+((B-1)*40/4);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
else
    error('FILE ERROR: potential file overwrite');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
%The AVSU is maybe in the wrong position.
end
function ldr(Q,B,X,I)
%Need to make ldr a re-executable function.
ii = DOT_data.index;
S = {X};
% For appearance of text- trying to make it look nice.
T = strcat(S, '               Mother');
Baby = replace(T,'Mother', 'Baby');
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{B},{[]},{[B]},{[]}, {[]}, {2*B}];
AA = [{B}, {[]}, {[]}, {[B]}, {[]}, {[B]}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:J7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'N7'); 
writecell(T, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
writecell(AA,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E8:J8');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C8');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D8');
writecell(Baby, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B8');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
mergeserver('C7:C8');
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10+[(B-1)*(6/4)];
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E8');
%For Entonox
q = 275+(B-1)*(6/2);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'I7');
%For Medical Air
q = 40+(B-1)*(40/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K8');
%For Vacuum
q = 40+(B-1)*(40/4);
q2 = 40;
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
writematrix(q2, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O8');
Eq = 130+[(B-1)*130/8];
AGS(I,B,Eq);
ii.index_prime = ii.index_prime+1;
mergeserver('T7:T8');
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{B}, {[]}, {[]}, {[]}, {B}, {[]}, {[]}, {[]}, {[]}, {[]}, {2*B}];
AA = [{[]}, {B}, {[]}, {[]}, {[]}, {[]}, {[]}, {B}, {[]},{[]}, {[]} {B}];
ii.index_prime = ii.index_prime+3;
np = num2str(ii.index_prime);
n2p = num2str(ii.index_prime+1);
R1 = ['E' np ':R' np]; R2 = ['E' n2p ':R' n2p ];
C1 = ['C' np]; C2 = ['C' n2p];
D1 = ['D' np]; D2 = ['D' n2p];
U1 = ['U' np]; 
BB1 = ['B' np]; BB2 = ['B' n2p];
CC = ['C' np ':C' n2p];
UU = ['U' np ':U' n2p];
E1 = ['E' np];
E2 = ['E' n2p];
I = ['I' np];
K = ['K' n2p];
O1 = ['O' np];
O2 = ['O' n2p];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R1);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C1);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D1);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', U1);
writecell(T, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB1);
writecell(AA,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range',R2);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C2);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D2);
writecell(Baby, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB2);
mergeserver(CC);
mergeserver(UU);
%For Oxygen
q = 10+[(B-1)*(6/4)];
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E1)
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E2)
%For Entonox
q = 275+(B-1)*(6/2);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', I)
%For Medical Air
q = 40+(B-1)*(40/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K)
%For Vacuum
q1 = 40+(B-1)*(40/4);
q2 = 40;
writematrix(q1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O1);
writematrix(q2, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O2);
Eq = 130+[(B-1)*130];
AGS(I,B,Eq);
ii.index_prime = ii.index_prime+1;
else
    error('FILE ERROR: potential file overwrite');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function lineac_bunker(Q,B,X,I)
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{B},{B},{[]},{B},{[]},{B},{B}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:K7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'N7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;         
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10 + [(B-1)*(6/3)];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Nitrous Oxide
q = 10 + [(B-1)*(6/4)];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'G7');
%For Medical Air
q = 40 + [(B-1)*40/4];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7')
%For Vacuum
q = 40 + [(B-1)*40/8];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
Eq = 130+[(B-1)*130/4];
AGS(I,B,Eq);
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{B},{[]},{[B]},{[]},{[]},{[]},{[B]},{[]},{[]},{[]},{B},{[]}, {B}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
UU = ['U' np];
BB = ['B' np];
E = ['E' np];
G = ['G' np];
K = ['K' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', UU);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);         
%For Oxygen
q = 10 + [(B-1)*(6/3)];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Nitrous Oxide
q = 10 + [(B-1)*(6/4)];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', G);
%For Medical Air
q = 40 + [(B-1)*40/4];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K);
%For Vacuum
q = 40 + [(B-1)*40/8];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
Eq = 130+[(B-1)*130/4];
AGS(I,B,Eq);
else
    error('FILE ERROR: potential file overwrite');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function MRI(Q,B,X,I)
%Need to make ldr a re-executable function.
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{B},{B},{[]},{B},{[]},{B},{B}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:K7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'N7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;         
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
%For Oxygen
q = 10 + [(B-1)*(6/3)];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Nitrous Oxide
q = 10 + [(B-1)*(6/4)];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'G7');
%For Medical Air
q = 40 + [(B-1)*40/4];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7')
%For Vacuum
q = 40 + [(B-1)*40/8];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{B},{[]},{[B]},{[]},{[]},{[]},{[B]},{[]},{[]},{[]},{B},{[]},{B}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
UU = ['U' np];
BB = ['B' np];
E = ['E' np];
G = ['G' np];
K = ['K' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', UU);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);         
%For Oxygen
q = 10 + [(B-1)*(6/3)];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Nitrous Oxide
q = 10 + [(B-1)*(6/4)];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', G);
%For Medical Air
q = 40 + [(B-1)*40/4];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K);
%For Vacuum
q = 40 + [(B-1)*40/8];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
Eq = 130+[(B-1)*130/4];
AGS(I,B,Eq);
else
    error('FILE ERROR: potential file overwrite');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function NICU(Q,B,X)
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{2*B},{[]},{[]},{2*B},{[]},{2*B},{[]}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:K7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(2, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'M7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q1 = 10 + (B-1)*6;
writematrix(q1,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Medical Air
q2 = 40.*B;
writematrix(q2,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7');
%For Vacuum
q3 = 40 + (B-1).*40/4;
writematrix(q3,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{2*B},{[]},{[]},{[]},{[]},{[]},{2*B},{[]},{[]},{[]},{2*B},{[]}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
TT = ['T' np];
BB = ['B' np];
E = ['E' np];
K = ['K' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(2, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', TT);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
%For Oxygen
q = 10 + [(B-1)*6];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Medical Air
q = 40*B;
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K);
%For Vacuum
q = 40 + [(B-1)*40/4];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
else
    error('FILE ERROR: potential file overwrite or open file');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function nursery(Q,B,X)
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{B},{[]},{[]},{[]},{[]},{B}, {[]}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:J7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'L7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10+((B-1)*(6/4));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Vacuum
q = 40;
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{B},{[]},{[]},{[]},{[]},{[]},{[]},{[]},{[]},{[]},{B},{[]}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
SS = ['S' np];
BB = ['B' np];
E = ['E' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', SS);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
%For Oxygen
q = 10+((B-1)*(6/4));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Vacuum
q = 40;
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
else
    error('FILE ERROR: potential file overwrite or open file');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function OT_gs(Q,B,X,I)
%Need to make ldr a re-executable function.
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{2*B},{2*B},{[]},{2*B},{[2*B]},{2*B}, {2*B}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:K7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E8:K8');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C8');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D8');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
mergeserver('C7:C8');
mergeserver('D7:D8');
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 100+[(B-1)*(10)];
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Nitrous Oxide
q = 15+(B-1)*6;
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'G7');
%For Medical Air
q = 40+(B-1)*(40/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7');
%For Surgical Air
if B > 4
qot = 350+(B-1)*350/4;
writematrix(qot, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'M7');
else
qot = 350+(B-1)*350/2;
writematrix(qot, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'M7');
end
%For Vacuum
q = 80+(B-1)*(80/2);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
Eq = 130+[(B-1)*130];
AGS(I,B,Eq);
ii.index_prime = ii.index_prime+1;
mergeserver('V7:V8');
mergeserver('Q7:Q8');
mergeserver('B7:B8');
mergeserver('E7:E8');
mergeserver('G7:G8');
mergeserver('K7:K8');
mergeserver('M7:M8');
mergeserver('O7:O8');
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{2*B},{[]},{2*B},{[]},{[]},{[]},{2*B},{[]},{[2*B]},{[]},{2*B},{[]},{2*B}];
ii.index_prime = ii.index_prime+3;
np = num2str(ii.index_prime);
n2p = num2str(ii.index_prime+1);
R1 = ['E' np ':R' np]; R2 = ['E' n2p ':R' n2p];
C1 = ['C' np]; C2 = ['C' n2p]; CC = ['C' np ':C' n2p];
D1 = ['D' np]; D2 = ['D' n2p]; DD = ['D' np ':D' n2p];
V = ['V' np]; VV = ['V' np ':V' n2p];
BB = ['B' np]; BbB = ['B' np ':B' n2p];
E1 = ['E' np]; Ee = ['E' np ':E' n2p];
G1 = ['G' np]; Gg = ['G' np ':G' n2p];
K1 = ['K' np]; Kk = ['K' np ':K' n2p];
M1 = ['M' np]; Mm = ['M' np ':M' n2p];
O1 = ['O' np]; Oo = ['O' np ':O' n2p]; 
QQ = ['Q' np ':Q' n2p];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R1);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C1);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D1);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', V);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range',R2);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C2);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D2);
%For Oxygen
q = 100+[(B-1)*(10)];
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E1);
%For Nitrous Oxide
q = 15+(B-1)*6;
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', G1);
%For Medical Air
q = 40+(B-1)*(40/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K1);
%For Surgical Air
if B > 4
qot = 350+(B-1)*350/4;
writematrix(qot, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', M1);
else
qot = 350+(B-1)*350/2;
writematrix(qot, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', M1);
end
%For Vacuum
q = 80+(B-1)*(80/2);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O1);
Eq = 130+[(B-1)*130];
AGS(I,B,Eq);
mergeserver(VV);
mergeserver(QQ);
mergeserver(CC);
mergeserver(DD);
mergeserver(BbB);
mergeserver(Ee);
mergeserver(Gg);
mergeserver(Kk);
mergeserver(Mm);
mergeserver(Oo);
ii.index_prime = ii.index_prime+1;
else
    error('FILE ERROR: potential file overwrite');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function OT_neurosurgery(Q,B,X,I)
%Need to make ldr a re-executable function.
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{2*B},{B},{[]},{2*B},{[2*B]},{2*B}, {B}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:K7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E8:K8');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C8');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D8');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
mergeserver('C7:C8');
mergeserver('D7:D8');
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 100+[(B-1)*(10)];
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Nitrous Oxide
q = 15+(B-1)*6;
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'G7');
%For Medical Air
q = 40+(B-1)*(40/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7');
%For Surgical Air
if B > 4
qot = 350+(B-1)*350/4;
writematrix(qot, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'M7');
else
qot = 350+(B-1)*350/2;
writematrix(qot, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'M7');
end
%For Vacuum
q = 80+(B-1)*(80/2);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
Eq = 130+[(B-1)*130];
AGS(I,B,Eq);
ii.index_prime = ii.index_prime+1;
mergeserver('V7:V8');
mergeserver('B7:B8');
mergeserver('E7:E8');
mergeserver('G7:G8');
mergeserver('K7:K8');
mergeserver('M7:M8');
mergeserver('O7:O8');
mergeserver('Q7:Q8');
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{2*B},{[]},{B},{[]},{[]},{[]},{2*B},{[]},{[2*B]},{[]},{2*B},{[]},{B}];
ii.index_prime = ii.index_prime+3;
np = num2str(ii.index_prime);
n2p = num2str(ii.index_prime+1);
R1 = ['E' np ':R' np]; R2 = ['E' n2p ':R' n2p];
C1 = ['C' np]; C2 = ['C' n2p]; CC = ['C' np ':C' n2p];
D1 = ['D' np]; D2 = ['D' n2p]; DD = ['D' np ':D' n2p];
V = ['V' np]; VV = ['V' np ':V' n2p];
BB = ['B' np]; BbB = ['B' np ':B' n2p];
E1 = ['E' np];  Ee = ['E' np ':E' n2p];
G1 = ['G' np];  Gg = ['G' np ':G' n2p];
K1 = ['K' np];  Kk = ['K' np ':K' n2p];
M1 = ['M' np];  Mm = ['M' np ':M' n2p];
O1 = ['O' np]; Oo = ['O' np ':O' n2p];
QQ = ['Q' np ':Q' n2p];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R1);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C1);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D1);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', V);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range',R2);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C2);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D2);
%For Oxygen
q = 100+[(B-1)*(10)];
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E1);
%For Nitrous Oxide
q = 15+(B-1)*6;
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', G1);
%For Medical Air
q = 40+(B-1)*(40/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K1);
%For Surgical Air
if B > 4
qot = 350+(B-1)*350/4;
writematrix(qot, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', M1);
else
qot = 350+(B-1)*350/2;
writematrix(qot, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', M1);
end
%For Vacuum
q = 80+(B-1)*(80/2);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O1);
Eq = 130+[(B-1)*130];
AGS(I,B,Eq);
mergeserver(VV);
mergeserver(QQ);
mergeserver(CC);
mergeserver(DD);
mergeserver(BbB);
mergeserver(Ee);
mergeserver(Gg);
mergeserver(Kk);
mergeserver(Mm);
mergeserver(Oo);
ii.index_prime = ii.index_prime+1;
else
    error('FILE ERROR: potential file overwrite');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function OT_orthopaedic(Q,B,X,I)
%Need to make ldr a re-executable function.
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{2*B},{B},{[]},{2*B},{[4*B]},{2*B}, {B}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:K7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E8:K8');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C8');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D8');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
mergeserver('C7:C8');
mergeserver('D7:D8');
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 100+[(B-1)*(10)];
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Nitrous Oxide
q = 15+(B-1)*6;
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'G7');
%For Medical Air
q = 40+(B-1)*(40/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7');
%For Surgical Air
if B > 4
qot = 350+(B-1)*350/4;
writematrix(qot, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'M7');
else
qot = 350+(B-1)*350/2;
writematrix(qot, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'M7');
end
%For Vacuum
q = 80+(B-1)*(80/2);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O8');
Eq = 130+[(B-1)*130];
AGS(I,B,Eq);
ii.index_prime = ii.index_prime+1;
mergeserver('V7:V8');
mergeserver('B7:B8');
mergeserver('E7:E8');
mergeserver('G7:G8');
mergeserver('K7:K8');
mergeserver('M7:M8');
mergeserver('O7:O8');
mergeserver('Q7:Q8');
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{2*B},{[]},{B},{[]},{[]},{[]},{2*B},{[]},{[4*B]},{[]},{2*B},{[]},{B}];
ii.index_prime = ii.index_prime+3;
np = num2str(ii.index_prime);
n2p = num2str(ii.index_prime+1);
R1 = ['E' np ':R' np]; R2 = ['E' n2p ':R' n2p];
C1 = ['C' np]; C2 = ['C' n2p]; CC = ['C' np ':C' n2p];
D1 = ['D' np]; D2 = ['D' n2p]; DD = ['D' np ':D' n2p];
V = ['V' np]; VV = ['V' np ':V' n2p];
BB = ['B' np]; BbB = ['B' np ':B' n2p];
E1 = ['E' np]; E2 = ['E' n2p]; Ee = ['E' np ':E' n2p];
G1 = ['G' np]; G2 = ['G' n2p]; Gg = ['G' np ':G' n2p];
K1 = ['K' np]; K2 = ['K' n2p]; Kk = ['K' np ':K' n2p];
M1 = ['M' np]; M2 = ['M' n2p]; Mm = ['M' np ':M' n2p];
O1 = ['O' np]; O2 = ['O' n2p]; Oo = ['O' np ':O' n2p];
QQ = ['Q' np ':Q' n2p];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R1);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C1);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D1);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', V);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range',R2);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C2);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D2);
%For Oxygen
q = 100+[(B-1)*(10)];
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E1);
%For Nitrous Oxide
q = 15+(B-1)*6;
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', G1);
%For Medical Air
q = 40+(B-1)*(40/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K1);
%For Surgical Air
if B > 4
qot = 350+(B-1)*350/4;
writematrix(qot, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', M1);
else
qot = 350+(B-1)*350/2;
writematrix(qot, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', M1);
end
%For Vacuum
q = 80+(B-1)*(80/2);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O1);
Eq = 130+[(B-1)*130];
AGS(I,B,Eq);
mergeserver(VV);
mergeserver(QQ);
mergeserver(CC);
mergeserver(DD);
mergeserver(BbB);
mergeserver(Ee);
mergeserver(Gg);
mergeserver(Kk);
mergeserver(Mm);
mergeserver(Oo);
ii.index_prime = ii.index_prime+1;
else
    error('FILE ERROR: potential file overwrite');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function par_emergency(Q,B,X)
%Need to make ldr a re-executable function.
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{2*B},{[]},{[]},{[2*B]},{[]},{2*B}, {[]}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:J7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(2, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'M7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10+((B-1)*(6/8));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Medical Air
q = 40+((B-1)*40/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7');
%For Vacuum
q = 40+((B-1)*40/4);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{2*B},{[]},{[]},{[]},{[]},{[]},{2*B},{[]},{[]},{[]},{2*B},{[]}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
TT = ['T' np];
BB = ['B' np];
E = ['E' np];
K = ['K' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(2, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', TT);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
%For Oxygen
q = 10+((B-1)*(6/8));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Medical Air
q = 40+((B-1)*40/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K);
%For Vacuum
q = 40+((B-1)*40/4);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
else
    error('FILE ERROR: potential file overwrite');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function par_ldr(Q,B,X)
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{B},{[]},{[]},{B},{[]},{B}, {[]}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:J7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(2, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'M7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10+((B-1)*(3/4));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Medical Air
q = 40+((B-1)*40/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7');
%For Vacuum
q = 40+((B-1)*40/4);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{B},{[]},{[]},{[]},{[]},{[]},{B},{[]},{[]},{[]},{B},{[]}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
TT = ['T' np];
BB = ['B' np];
E = ['E' np];
K = ['K' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(2, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', TT);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
%For Oxygen
q = 10+((B-1)*(3/4));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Medical Air
q = 40+((B-1)*40/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K);
%For Vacuum
q = 40+((B-1)*40/4);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
else
    error('FILE ERROR: potential file overwrite');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function par_mental(Q,B,X)
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{B},{[]},{[]},{B},{[]},{B}, {[]}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:J7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(2, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'M7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10+((B-1)*(6/4));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Medical Air
q = 40+((B-1)*40/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7');
%For Vacuum
q = 40+((B-1)*40/4);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{B},{[]},{[]},{[]},{[]},{[]},{B},{[]},{[]},{[]},{B},{[]}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
TT = ['T' np];
BB = ['B' np];
E = ['E' np];
K = ['K' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(2, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', TT);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
%For Oxygen
q = 10+((B-1)*(6/4));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Medical Air
q = 40+((B-1)*40/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K);
%For Vacuum
q = 40+((B-1)*40/4);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
else
    error('FILE ERROR: potential file overwrite');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
%The AVSU is maybe in the wrong position.
end
function par_ot(Q,B,X)
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{2*B},{[]},{[]},{2*B},{[]},{2*B}, {[]}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:J7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(2, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'M7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10+((B-1)*(6));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Medical Air
q = 40+((B-1)*10/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7');
%For Vacuum
q = 40+((B-1)*40/4);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{2*B},{[]},{[]},{[]},{[]},{[]},{2*B},{[]},{[]},{[]},{2*B},{[]}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
TT = ['T' np];
BB = ['B' np];
E = ['E' np];
K = ['K' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(2, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', TT);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
%For Oxygen
q = 10+((B-1)*(6));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Medical Air
q = 40+((B-1)*10/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K);
%For Vacuum
q = 40+((B-1)*40/4);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
else
    error('FILE ERROR: potential file overwrite');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function Plaster_room(Q,B,X,I)
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{B},{B},{B},{B},{B},{B}, {B}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:K7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'P7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10+[(B-1)*(6/4)];
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7')
%For Nitrogen
q = 15+(B-1)*(6);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'G7')
%For Entonox
q = 20+(B-1)*(10/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'I7')
%For Medical Air
q = 40+(B-1)*(20/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7')
%For Surgical Air
q = 350;
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'M7');
%For Vacuum
q = 40+(B-1)*(40/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
Eq = 130+[(B-1)*130/8];
AGS(I,B,Eq);
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{B},{[]},{B},{[]},{B},{[]},{B},{[]},{B},{[]},{B},{[]},{B}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
BB = ['B' np];
E = ['E' np];
G = ['G' np];
I = ['I' np];
K = ['K' np];
M = ['M' np];
O = ['O' np];
W = ['W' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', W);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
%For Oxygen
q = 10+(B-1)*(6/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Nitrogen
q = 15+(B-1)*(6);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', G);
%For Entonox
q = 20+(B-1)*(10/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', I);
%For Medical Air
q = 40+(B-1)*(20/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K);
%For Surgical Air
q = 350;
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', M)
%For Vacuum
q = 40+(B-1)*(40/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
Eq = 130+[(B-1)*130/8];
AGS(I,B,Eq);
else
    error('FILE ERROR: potential file overwrite or open file');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function renal(Q,B,X)
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{B},{[]},{[]},{B},{[]},{B}, {[]}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:J7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'M7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10+((B-1)*(6/4));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%Medical Air
q = 20+((B-1)*(10/4));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7');
%For Vacuum
q = 40+((B-1)*40/4);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{B},{[]},{[]},{[]},{[]},{[]},{B},{[]},{[]},{[]},{B},{[]}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
TT = ['T' np];
BB = ['B' np];
E = ['E' np];
K = ['K' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', TT);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
%For Oxygen
q = 10+((B-1)*(6/4));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Medical Air
q = 20+((B-1)*(10/4));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K);
%For Vacuum
q = 40 +((B-1)*(40/4));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
else
    error('FILE ERROR: potential file overwrite or open file');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
%The AVSU is maybe in the wrong position.
end
function Resuscitation_Room(Q,B,X,I)
%Need to make ldr a re-executable function.
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{2*B},{[2*B]},{[]},{[2*B]},{[]},{[2*B]}, {[2*B]}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:K7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(2, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'N7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;         
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
%For Oxygen
q = 100 + [(B-1)*(6/4)];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Nitrous Oxide
q = 10 + [(B-1)*(6/4)];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'G7');
%For Medical Air
q = 40 + [(B-1)*20/4];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7')
%For Vacuum
q = 40 + [(B-1)*40/4];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7')
Eq = 130+(B-1)*(130/4);
AGS(I,B,Eq);
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{2*B},{[]},{[2*B]},{[]},{[]},{[]},{[2*B]},{[]},{[]},{[]},{2*B},{[]},{2*B}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
UU = ['U' np];
BB = ['B' np];
E = ['E' np];
G = ['G' np];
K = ['K' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(2, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', UU);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);         
%For Oxygen
q = 100 + [(B-1)*(6/4)];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Nitrous Oxide
q = 10 + [(B-1)*(6/4)];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', G);
%For Medical Air
q = 40 + [(B-1)*20/4];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K);
%For Vacuum
q = 40 + [(B-1)*40/4];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
Eq = 130+(B-1)*(130/4);
AGS(I,B,Eq);
else
    error('FILE ERROR: potential file overwrite');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function special_procedures(Q,B,X,I)
%Need to make ldr a re-executable function.
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{B},{B},{[]},{B},{[]},{B},{B}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:K7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'N7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;         
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10 + [(B-1)*(6/3)];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Nitrous Oxide
q = 10 + [(B-1)*(6/4)];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'G7');
%For Medical Air
q = 40 + [(B-1)*40/4];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7')
%For Vacuum
q = 40 + [(B-1)*40/8];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
Eq = 130+[(B-1)*130/8];
AGS(I,B,Eq);
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{B},{[]},{[B]},{[]},{[]},{[]},{[B]},{[]},{[]},{[]},{B},{[]},{B}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
UU = ['U' np];
BB = ['B' np];
E = ['E' np];
G = ['G' np];
K = ['K' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', UU);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);         
%For Oxygen
q = 10 + [(B-1)*(6/3)];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Nitrous Oxide
q = 10 + [(B-1)*(6/4)];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', G);
%For Medical Air
q = 40 + [(B-1)*40/4];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K);
%For Vacuum
q = 40 + [(B-1)*40/8];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
Eq = 130+[(B-1)*130/8];
AGS(I,B,Eq);
else
    error('FILE ERROR: potential file overwrite');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.

end
function Surgeon_aadca(Q,B,X)
%Need to make ldr a re-executable function.
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{B},{B},{[]},{B},{[]},{B},{B}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:K7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'N7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;         
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10 + [(B-1)*(6/3)];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Nitrous Oxide
q = 10 + [(B-1)*(6/4)];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'G7');
%For Medical Air
q = 40 + [(B-1)*40/4];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7')
%For Vacuum
q = 40 + [(B-1)*40/8];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{B},{[]},{[B]},{[]},{[]},{[]},{[B]},{[]},{[]},{[]},{B},{B}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
UU = ['U' np];
BB = ['B' np];
E = ['E' np];
G = ['G' np];
K = ['K' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', UU);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);         
%For Oxygen
q = 10 + [(B-1)*(6/3)];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Nitrous Oxide
q = 10 + [(B-1)*(6/4)];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', G);
%For Medical Air
q = 40 + [(B-1)*40/4];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K);
%For Vacuum
q = 40 + [(B-1)*40/8];
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
else
    error('FILE ERROR: potential file overwrite');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function tomography (Q,B,X,I)
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{B},{[]},{[]},{[]},{[]},{B}, {[]}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:J7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'L7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10+((B-1)*(6/3));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Vacuum
q = 40+((B-1)*40/8);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
Eq = 130+[(B-1)*130/4];
AGS(I,B,Eq);
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{B},{[]},{[]},{[]},{[]},{[]},{[]},{[]},{[]},{[]},{B},{[]}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
SS = ['S' np];
BB = ['B' np];
E = ['E' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', SS);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
%For Oxygen
q = 10+((B-1)*(6/3));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Vacuum
q = 40 +((B-1)*(40/8));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
Eq = 130+[(B-1)*130/4];
AGS(I,B,Eq);
else
    error('FILE ERROR: potential file overwrite or open file');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function treatment_room(Q,B,X)
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{B},{[]},{[]},{B},{[]},{B}, {[]}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:J7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'M7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10+((B-1)*(6/4));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Medical Air
q = 20+((B-1)*10/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7');
%For Vacuum
q = 40+((B-1)*40/8);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{B},{[]},{[]},{[]},{[]},{[]},{B},{[]},{[]},{[]},{B},{[]}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
TT = ['T' np];
BB = ['B' np];
E = ['E' np];
K = ['K' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', TT);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
%For Oxygen
q = 10+((B-1)*(6/4));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Medical Air
q = 20+((B-1)*10/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', K);
%For Vacuum
q = 40+((B-1)*40/8);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
else
    error('FILE ERROR: potential file overwrite');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
%The AVSU is maybe in the wrong position.
end
function ultrasound (Q,B,X,I)
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{B},{[]},{[]},{[]},{[]},{B}, {[]}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:J7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'L7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10+((B-1)*(6/3));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Vacuum
q = 40+((B-1)*40/8);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
Eq = 130+[(B-1)*130/4];
AGS(I,B,Eq);
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{B},{[]},{[]},{[]},{[]},{[]},{[]},{[]},{[]},{[]},{B},{[]}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
SS = ['S' np];
BB = ['B' np];
E = ['E' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', SS);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
%For Oxygen
q = 10+((B-1)*(6/3));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Vacuum
q = 40 +((B-1)*(40/8));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
Eq = 130+[(B-1)*130/4];
AGS(I,B,Eq);
else
    error('FILE ERROR: potential file overwrite or open file');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function urography (Q,B,X,I)
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{B},{[]},{[]},{[]},{[]},{B}, {[]}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:J7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'L7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10+((B-1)*(6/3));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Vacuum
q = 40+((B-1)*40/8);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
Eq = 130+[(B-1)*130/4];
AGS(I,B,Eq);
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{B},{[]},{[]},{[]},{[]},{[]},{[]},{[]},{[]},{[]},{B},{[]}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
SS = ['S' np];
BB = ['B' np];
E = ['E' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', SS);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
%For Oxygen
q = 10+((B-1)*(6/3));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Vacuum
q = 40 +((B-1)*(40/8));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
Eq = 130+[(B-1)*130/4];
AGS(I,B,Eq);
else
    error('FILE ERROR: potential file overwrite or open file');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function Wash_room(Q,B,X)
%Need to make ldr a re-executable function.
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{[]},{[]},{[]},{Q*B},{[Q*B]},{[]}, {[]}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:K7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'N7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{[]},{[]},{Q*B},{[Q*B]},{[]}, {[]}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':K' np];
C = ['C' np];
D = ['D' np];
N = ['N' np];
BB = ['B' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', N);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
else
    error('FILE ERROR: potential file overwrite or open file');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end
function x_ray_room(Q,B,X,I)
ii = DOT_data.index;
S = {X};
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{B},{[]},{[]},{[]},{[]},{B}, {[]}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:J7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'L7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2double(n);
ii.index_prime = nc;
Add_columns_flow ('E4'); 
Add_columns_flow('G4');
Add_columns_flow('I4');
Add_columns_flow('K4');
Add_columns_flow('M4');
Add_columns_flow('O4');
Add_columns_flow('Q6');
%For Oxygen
q = 10+((B-1)*(6/3));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
%For Vacuum
q = 40+((B-1)*40/8);
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
Eq = 130+[(B-1)*130/4];
AGS(I,B,Eq);
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{[]},{B},{[]},{[]},{[]},{[]},{[]},{[]},{[]},{[]},{[]},{B},{[]}];
ii.index_prime = ii.index_prime+1;
np = num2str(ii.index_prime);
R = ['E' np ':R' np];
C = ['C' np];
D = ['D' np];
SS = ['S' np];
BB = ['B' np];
E = ['E' np];
O = ['O' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', SS);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB);
%For Oxygen
q = 10+((B-1)*(6/3));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', E);
%For Vacuum
q = 40 +((B-1)*(40/8));
writematrix(q,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', O);
Eq = 130+[(B-1)*130/4];
AGS(I,B,Eq);
else
    error('FILE ERROR: potential file overwrite or open file');
end
end
 end
end
