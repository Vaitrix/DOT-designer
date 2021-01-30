function OT_g(Q,B,X)
%Need to make ldr a re-executable function.
ii = DOT_data.index;
S = {X};
% Generates a workable spreadsheet from the template which is preserved.
copyfile('DOT-table-Blankslate.xlsx', 'DOT(2)-table-Blankslate.xlsx');
movefile('DOT(2)-table-Blankslate.xlsx', 'modular-DOT-Table.xlsx');
% Writes data from MATLAB to Excel spreadsheets
A = [{2*Q*B},{2*Q*B},{[]},{2*Q*B},{[2*Q*B]},{2*Q*B}, {2*Q*B}];
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
%For Oxygen
q = 100+[(B-1)*(10)];
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7');
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E8');
%For Nitrous Oxide
q = 15+(B-1)*6;
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'G7');
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'G8');
%For Medical Air
q = 40+(B-1)*(40/4);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K7');
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'K8');
%For Surgical Air
if B > 4
qot = 350+(B-1)*350/4;
writematrix(qot, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'M7');
writematrix(qot, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'M8');
else
qot = 350+(B-1)*350/2;
writematrix(qot, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'M7');
writematrix(qot, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'M8');
end
%For Vacuum
q = 80+(B-1)*(80/2);
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O7');
writematrix(q, 'modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'O8');
ii.index_prime = ii.index_prime+1;
mergeserver('U7:U8');
end