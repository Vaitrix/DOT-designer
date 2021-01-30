function renal(Q,B,X)
%Need to make ldr a re-executable function.
persistent ii
S = {X};
if isempty(ii) == 1
    ii = 0;
end
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{Q*B},{[]},{[]},{Q*B},{[]},{Q*B}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:J7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'M7'); 
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
n = '7';
nc = str2num(n);
ii = nc;
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{Q*B},{[]},{[]},{Q*B},{[]},{Q*B}];
ii = ii+1;
np = num2str(ii);
R = ['E' np ':J' np];
C = ['C' np];
D = ['D' np];
M = ['M' np];
BB = ['B' np];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', M);
writecell(S, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB)
else
    error('FILE ERROR: potential file overwrite');
    clear ldr
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
%The AVSU is maybe in the wrong position.
end