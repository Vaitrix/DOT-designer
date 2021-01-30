function S = SaveAs(X)
[file, name, ~] = uiputfile('*.xlsx', 'Save file name', '');
          [~, y] = size(name);
          n2 = name(1, 1:(y-1));
         S = '';
         a = ['~$' file];
if isfile('~$modular-DOT-Table.xlsx') == 0 && isfile(a) == 0 
          Floor_name(X);
          column_resize;
          if  strcmp(n2, cd) == 1
          copyfile('modular-DOT-Table.xlsx', 'modular(2)-DOT-Table.xlsx');
          movefile('modular-DOT-Table.xlsx', file);
          movefile('modular(2)-DOT-Table.xlsx', 'modular-DOT-Table.xlsx');
          delete('modular(2)-DOT-Table.xlsx');
          elseif strcmp(n2, cd) == 0
          copyfile('modular-DOT-Table.xlsx', 'modular(2)-DOT-Table.xlsx');
          movefile('modular-DOT-Table.xlsx', file);
          movefile(file, name, 'f');
          movefile('modular(2)-DOT-Table.xlsx', 'modular-DOT-Table.xlsx');
          delete('modular(2)-DOT-Table.xlsx');
          else 
            S = 'Close the file to be overwritten';
          end
else 
             S = 'Close the file to be saved';
end
end
