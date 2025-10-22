function [tableOut,badRecordTable] = readAndCombineXlsxRecord(xlsxFileName,keyColumnHeaders)
% given: a xls table path, header names, and a save file location
% DO: step through each sheet, pull out the column based on first header
% location, and combine them into a master table and return that table

% example input
% xlsxFileName = 'Z:\PearceLabRecords\Mouse Inventory\Lamp5-cre\Lamp5-cre.xlsx'; % we'll step through a list of these, but start here for testing
% keyColumnHeaders = {'ID Number','DOB','Date of Exp','mouseAssignment','sacCode','fundingID'};

% toggle verbose mode
showBadRecords = true;
badRecordTable = table;

tableOut = table;
sheetList = sheetnames(xlsxFileName);
for iSheet = 1:size(sheetList,1)
    foundData = false;
    opts = detectImportOptions(xlsxFileName,'Sheet',sheetList{iSheet});
    opts.VariableNamingRule = 'preserve';
    % we ran into an issue searching for the first correct row, but reading
    % it in as a cell array and finding the earliest instance across all
    % columns (even if some are missing) seems to work
    rawSheetAsCell = readcell(xlsxFileName, 'Sheet', sheetList{iSheet});
    earliestHeaderRow = inf; 
    for iCol = 1:size(rawSheetAsCell, 2)
        columnData = string(rawSheetAsCell(:, iCol));
        for iHeader = 1:size(keyColumnHeaders, 2)
            targetHeader = string(keyColumnHeaders{iHeader});
            singleColumn = (columnData == targetHeader);
            if any(singleColumn) 
                foundRow = find(singleColumn, 1, 'first');
                earliestHeaderRow = min(earliestHeaderRow, foundRow);
            end
        end
    end
    % opts.DataLines = [earliestHeaderRow + 1, inf];
    colNumber = size(rawSheetAsCell,2);
    remainder = mod(colNumber - 1, 26);
    % Convert the remainder (0-25) to an ASCII character (A=65)
    colLetter = char('A' + remainder);
    opts.DataRange = ['A', num2str(earliestHeaderRow), ':', colLetter, num2str(size(columnData,1))]; 
    
    % here, we'll fix all imported rows as 'char' so we can search it easily.
    numColumns = size(rawSheetAsCell,2);
    fixedType = {'char'};
    charArray = repmat(fixedType, 1, numColumns);
    opts.VariableTypes(1:numColumns) = charArray;
    thisTable = readtable(xlsxFileName,opts); 

    % TODO: figure out how to load it once only
    % fixedType = {'char'};
    % numColumns = size(thisTable,2);
    % charArray = repmat(fixedType, 1, numColumns);
    % opts.VariableTypes(1:numColumns) = charArray;
    % opts.VariableNamingRule = 'preserve'; 

    % we could use opt.VariableNames to confirm we got the right columns?


    % Now search through and create a sub table of specified var from this sheet
    tablefromSheetOut = table('Size',[size(columnData,1),length(keyColumnHeaders)],'VariableTypes',{'string','string','string','string','string','string'},'VariableNames',keyColumnHeaders);
    % thisTable = readtable(xlsxFileName,opts);
    for iHeader = 1:size(keyColumnHeaders,2)
        % step through each column and look for matching key words until we
        % find what we're looking for (or run out of columns)
        foundData = false;
        
        tempTable = table;
        iCol = 1;
        while foundData == false 
            singleColumn = ismember(thisTable{:,iCol},keyColumnHeaders{iHeader});
            if sum(singleColumn)>0
                thisRow = find(singleColumn,1,'first');
                tempTable = thisTable(thisRow+1:end,iCol);
                tempTable = renamevars(tempTable, tempTable.Properties.VariableNames, keyColumnHeaders(iHeader));
                foundData = true;
                doOnce = true;
            end
            if iCol >= size(thisTable,2)
                foundData = true;
            end
            iCol = iCol+1;
        end
        if foundData & ~isempty(tempTable)
            % tablefromSheetOut = horzcat(tablefromSheetOut, tempTable);
            tablefromSheetOut.(keyColumnHeaders{iHeader})(1:size(tempTable.(keyColumnHeaders{iHeader}),1)) = tempTable.(keyColumnHeaders{iHeader});
        end
    end
    tablefromSheetOutClean = table('Size',[height(tablefromSheetOut),length(keyColumnHeaders)],'VariableTypes',{'string','string','string','string','string','string'},'VariableNames',keyColumnHeaders);
    for iiHeader = 1:length(keyColumnHeaders)
        if ismember(keyColumnHeaders{iiHeader}, tablefromSheetOut.Properties.VariableNames)
            tablefromSheetOutClean.(keyColumnHeaders{iiHeader}) = tablefromSheetOut.(keyColumnHeaders{iiHeader});
        end
    end
    if foundData
        % let's look for correct dates and exclude anything without a good date
        datetimeFormatString = 'dd-MMM-yyyy';
        tablefromSheetOutClean.("DOB") = datetime(tablefromSheetOutClean.("DOB"), 'InputFormat', datetimeFormatString);
        missingDates = isnat(tablefromSheetOutClean.("DOB"));
        badRecordTable = vertcat(tablefromSheetOutClean(missingDates,:), badRecordTable);
        if showBadRecords
            disp(['File: ' xlsxFileName ' Sheet: ' sheetList{iSheet} ' excluded ' num2str(sum(missingDates)) ' records due to bad date formatting.']);
        end
        tablefromSheetOutClean(missingDates,:) = [];
        tableOut = vertcat(tableOut, tablefromSheetOutClean);
    end

end


% many columns still have missing data. working on it.



% TODO!!  noticing a problem where DOB is entered below and some records
% are correctly included and others are not.  THIS NEEDS TO BE SOLVED but I
% ran out of time today 10/10/25.  I think we can just check for both ID
% and DOB for any row, and include it if they both exist?  try that and see
% if that helps.  and no need to include totally empty rows in the excluded
% records (they aren't even records, just bad input).
