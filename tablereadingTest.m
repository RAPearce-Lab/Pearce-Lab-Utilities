% **** outline of mouse table curating we need to perform ****
% 1. load in the "table of tables" so we know what to operate on
% 2. for each table, load it in with appropriate opts - this is proving
%   challenging because columns might be formatted as doubles or whatever.
% 3. on the find the specific columns of interest, in case there's trouble (they don't all start at the same row)
% 4. load in the specific columns and create the counts
% ... do we want to edit and rewrite from here?

%
tableTest = 'Z:\PearceLabRecords\Mouse Inventory\Lamp5-cre\Lamp5-cre.xlsx'; % we'll step through a list of these, but start here for testing
keyColumnHeaders = {'mouseAssignment','sacCode','fundingID'};
% warning!  this is set to find the last instance of these "keyColumnHeaders" and base the new table around that.  consider this if the table output is weird, and I will consider a better way. 
saveFileName = 'Z:\PearceLabRecords\Mouse Inventory\Lamp5-cre\mouseCount.xlsx';
% === BEGIN ===
% given:  a single table,sheet
% DO: step through each sheet, and return the row (and column) location of the start of the
% mouse data.
fullAnimalTable = table;
sheetList = sheetnames(tableTest);
for i = 1:size(sheetList,1)
    foundData = false;
    opts = detectImportOptions(tableTest,'Sheet',sheetList{i});
    % here, we'll fix all imported rows as 'char' so we can search it easily.
    thisTable = readtable(tableTest,opts); % unfortunately, we need to load this twice - first to see how many columns we have
    fixedType = {'char'};
    numColumns = size(thisTable,2);
    charArray = repmat(fixedType, 1, numColumns);
    opts.VariableTypes(1:numColumns) = charArray;
    % Now search through and create a sub table of specified var from this sheet
    tempTable = table;
    tablefromSheetOut = table;
    thisTable = readtable(tableTest,opts);
    for iHeader = 1:size(keyColumnHeaders,2)
        for iCol = 1:size(thisTable,2)
            singleColumn = ismember(thisTable{:,iCol},keyColumnHeaders{iHeader});
            if sum(singleColumn)>0
                thisRow = find(singleColumn,1,'first');
                % from here we can create a huge table of all mice and codes, or
                % just get the numbers we want (and will go with the latter for now)
                tempTable = thisTable(thisRow+1:end,iCol);
                foundData = true;
            end
        end
        if foundData
            tablefromSheetOut = horzcat(tablefromSheetOut, tempTable);
        end
    end
    if foundData
        tablefromSheetOut = horzcat(thisTable(thisRow+1:end,1), tablefromSheetOut);
        oldNames = tablefromSheetOut.Properties.VariableNames;
        newNames = {'ID Number','mouseAssignment','sacCode','fundingID'};
        tablefromSheetOut = renamevars(tablefromSheetOut, oldNames, newNames);
        fullAnimalTable = vertcat(fullAnimalTable, tablefromSheetOut);
    end
end
% quick cleanup.  some of the rows we found (like the ones at the end or
% ones we didn't enter) are empty and should be removed
eliminateThese = ismissing(fullAnimalTable.("mouseAssignment"));
fullAnimalTable(eliminateThese,:) = [];

writetable(fullAnimalTable,saveFileName);
