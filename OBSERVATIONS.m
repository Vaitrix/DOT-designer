%%OBESERVATIONS
%When creating a new function, applying the function to an already existing
%Excel spreadsheet leads to an error- and this can be explained via the
%persistent variable ii being a blank array- this would have to be
%accounted for when developing the app- a way to make the persistent
%variable ii more global throughout the entire program. (not resolved)
%If I open the Excel document and then start inputting data into the
%variables. The function gets arrested- so that's a bug that needs fixing.
%Check whether the terminal outlets for the OT are right.
%Included a nested if statement in Tomography.m which would overwrite the
%bug involved in using another function to write a file.
%Confirm whether a He/O2 column needs to be generated (only for IAP room)
%Did everything except the oral surgery departments and a few surgical
%departments.

% Write some functions that automatically save the Excel file when the app
% is closed.
%Check out the error prompts- how to fix that within the program- final
%call.
%Checks Burns Unit flow calculation.
%Confirm how many gases for AVSU.-Recheck all AVSUS.
%General Purpose Rooms is the same as an inpatient room. 
%Do the Operation Theatres later along with the LDR rooms.
%Check out Plaster rooms for nitrous oxide flow.
%AGS Venturi required for Entonox.- LDR Cluster.
%Anaesthetist aadca room flow calculations incomplete
%Complete the IAP room for flow (if there is any).
%DO LDR ROOM FOR FLOW AND OTHER CORRECTIONS.
%Adding a room when excel file is open- need an error prompt.
%ICU rooms- make it 4x rather than 2x.
%C-section OT need to be added into the software.
%No ENT for Plaster Room.
% Change Operating Theatre- fuse flow calculations.
% For all rooms providing AGS Venturi- will provide 50 litres per minute. 
% For AGSS- another equation- ...........- check HTM 02-01.
% Problems with using multiple inputs in a function- what if one of them
% isnt used.
%Calculations for Anaesthetist room....



%% function CONDENSE
%Condense is dependent upon the index prime as its memory cache and so can
%only be integrated into the DOT table- can't use the function outside of
%the DOT app and it has to be utilized at the very end.

