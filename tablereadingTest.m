% **** outline of mouse table curating we need to perform ****
% 1. load in the "table of tables" so we know what to operate on
% 2. for each table, load it in with appropriate opts - this is proving
%   challenging because columns might be formatted as doubles or whatever.
% 3. on the find the specific columns of interest, in case there's trouble (they don't all start at the same row)
% 4. load in the specific columns and create the counts
% ... do we want to edit and rewrite from here?

%
xlsxFileName = 'Z:\PearceLabRecords\Mouse Inventory\Lamp5-cre\Lamp5-cre.xlsx'; % we'll step through a list of these, but start here for testing
keyColumnHeaders = {'ID Number','DOB','Date of Exp','mouseAssignment','sacCode','fundingID'};
saveFileName = 'Z:\PearceLabRecords\Mouse Inventory\Lamp5-cre\mouseCount.xlsx';


% our new function
[fullAnimalTable] = readAndCombineXlsxRecord(xlsxFileName,keyColumnHeaders);

% cleanup here?
% some of the rows we found (like the ones at the end or
% ones we didn't enter) are empty and should be removed
% eliminateThese = ismissing(fullAnimalTable.("mouseAssignment"));
%fullAnimalTable(eliminateThese,:) = [];

writetable(fullAnimalTable,saveFileName);


