function AGS(I,B,W)
if strcmp(I, 'AGS Venturi')
 ii = DOT_data.index;
 X = ii.index_prime;
 S = num2str(X);
 G = W;
 Q = ['Q' S];
 writematrix(G,'modular-DOT-Table.xlsx', 'Sheet', 1,'Range', Q); 
elseif strcmp(I, 'AGSS')
 ii = DOT_data.index;
 X = ii.index_prime;
 S = num2str(X);
 G = 50*B;
 Q = ['Q' S];
 writematrix(G,'modular-DOT-Table.xlsx', 'Sheet', 1,'Range', Q);
else   
end
end

