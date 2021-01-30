classdef flow
   
    properties
    end
    
    methods (Static)
        function anaesthetic_room (B)
            B = [];
            Add_columns_flow('E4');
            
        end
        function Anaesthetist_aadca
            % For calculating oxygen flow.
            Add_columns_flow ('E4');
            
        end
        function angiography  
           ii = DOT_data.index;
           y = num2str(ii.index_prime);           
           Add_columns_flow ('E4'); 
           A = readmatrix('modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E7:E7');
           B = A(1);
           Q = 10+[(B-1)*(6/3)];
           writematrix('modular-DOT-Table.xlsx', 'Sheet', 1, 'Range', 'E');
           
        end
        
    end
    
end
