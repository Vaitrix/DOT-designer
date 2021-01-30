function ldr(Q,B,X)
%Need to make ldr a re-executable function.
ii = DOT_data_trial.index1;
S = {X};
% For appearance of text- trying to make it look nice.
T = strcat(S, '               Mother');
Baby = replace(T,'Mother', 'Baby');
if isfile('modular-DOT-Table.xlsx') == 0
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{Q*B},{[]},{[Q*B]},{[]}, {[]}, {2*Q*B}];
AA = [{Q*B}, {[]}, {[]}, {[Q*B]}, {[]}, {[Q*B]}];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E7:J7');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C7');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D7');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'N7'); 
writecell(T, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B7');
writecell(AA,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range','E8:J8');
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'C8');
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'D8');
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'N8'); 
writecell(Baby, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'B8');
n = '7';
nc = str2double(n);
ii.index1_prime = nc;
mergeserver('C7:C8');
mergeserver('N7:N8');
elseif isfile('modular-DOT-Table.xlsx') == 1 && ii.index1_prime >= 7 && isfile('~$modular-DOT-Table.xlsx') == 0
A = [{2*Q*B},{Q*B},{[]},{2*Q*B},{[]},{2*Q*B}];
AA = [{Q*B}, {[]}, {[]}, {[Q*B]}, {[]}, {[Q*B]}];
ii.index1_prime = ii.index1_prime+2;
np = num2str(ii.index1_prime);
n2p = num2str(ii.index1_prime+1);
R1 = ['E' np ':J' np]; R2 = ['E' n2p ':J' n2p ];
C1 = ['C' np]; C2 = ['C' n2p];
D1 = ['D' np]; D2 = ['D' n2p];
N1 = ['N' np]; N2 = ['N' n2p];
BB1 = ['B' np]; BB2 = ['B' n2p];
CC = ['C' np ':C' n2p];
NN = ['N' np ':N' n2p];
writecell(A,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', R1);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C1);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D1);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', N1);
writecell(T, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB1);
writecell(AA,'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range',R2);
writematrix(Q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', C2);
writematrix(B, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', D2);
writematrix(1, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', N2); 
writecell(Baby, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', BB2);
mergeserver(CC);
mergeserver(NN);
else
    error('FILE ERROR: potential file overwrite');
end
%Need to change code in order to ensure LDR function takes over from where
%the other rooms left of.
%Need to resize the  columns after the change is executed via writematrix.
%Need a way to reset the function at each floor.
%Highlight the rows of each floor.
end








