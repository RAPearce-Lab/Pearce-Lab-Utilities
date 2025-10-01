% **** outline of mouse table curating we need to perform ****
% 1. load in the "table of tables" so we know what to operate on
% 2. for each table, load it in with appropriate opts - this is proving
%   challenging because columns might be formatted as doubles or whatever.
% 3. on the find the specific columns of interest, in case there's trouble (they don't all start at the same row)
% 4. load in the specific columns and create the counts
% ... do we want to edit and rewrite from here?

% first, some testing
tableTest = 'Z:\PearceLabRecords\Mouse Inventory\Lamp5-cre\Lamp5-cre.xlsx';
keyColumnHeaders = {'mouseAssignment','sacCode','fundingID'};
% given:  a single table
% DO: step through each sheet, and return the row (and column) location of the start of the
% mouse data.

sheetList = sheetnames(tableTest);
i = 4; % this will be a loop later.  starting with a page I know
opts = detectImportOptions(tableTest,'Sheet',sheetList{i});

fixedType = {'char'};
numColumns = size(thisTable,2);
charArray = repmat(fixedType, 1, numColumns);
opts.VariableTypes(1:numColumns) = charArray;

thisTable = readtable(tableTest,opts);
for iCol = 1:size(thisTable,2)
    singleColumn = ismember(thisTable{:,iCol},keyColumnHeaders{1});
    if sum(singleColumn)>0
        thisRow = find(singleColumn,1,'first');
        thisColumn = iCol;
    end
end



