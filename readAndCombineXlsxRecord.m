function [tableOut] = readAndCombineXlsxRecord(xlsxFileName,keyColumnHeaders)
% given: a xls table path, header names, and a save file location
% DO: step through each sheet, pull out the column based on first header
% location, and combine them into a master table and return that table

% example input
% xlsxFileName = 'Z:\PearceLabRecords\Mouse Inventory\Lamp5-cre\Lamp5-cre.xlsx'; % we'll step through a list of these, but start here for testing
% keyColumnHeaders = {'ID Number','DOB','Date of Exp','mouseAssignment','sacCode','fundingID'};

tableOut = table;
sheetList = sheetnames(xlsxFileName);
for iSheet = 1:size(sheetList,1)
    foundData = false;
    opts = detectImportOptions(xlsxFileName,'Sheet',sheetList{iSheet});
    opts.VariableNamingRule = 'preserve'; 
    % here, we'll fix all imported rows as 'char' so we can search it easily.
    thisTable = readtable(xlsxFileName,opts); % unfortunately, we need to load this twice - first to see how many columns we have
    % TODO: figure out how to load it once only
    fixedType = {'char'};
    numColumns = size(thisTable,2);
    charArray = repmat(fixedType, 1, numColumns);
    opts.VariableTypes(1:numColumns) = charArray;
    opts.VariableNamingRule = 'preserve'; 
    % Now search through and create a sub table of specified var from this sheet
    tablefromSheetOut = table;
    thisTable = readtable(xlsxFileName,opts);
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
            end
            if iCol >= size(thisTable,2)
                foundData = true;
            end
            iCol = iCol+1;
        end
        if foundData
            tablefromSheetOut = horzcat(tablefromSheetOut, tempTable);
        end
    end
    tablefromSheetOutClean = table('Size',[height(tablefromSheetOut),length(keyColumnHeaders)],'VariableTypes',{'string','string','string','string','string','string'},'VariableNames',keyColumnHeaders);
    for iiHeader = 1:length(keyColumnHeaders)
        if ismember(keyColumnHeaders{iiHeader}, tablefromSheetOut.Properties.VariableNames)
            tablefromSheetOutClean.(keyColumnHeaders{iiHeader}) = tablefromSheetOut.(keyColumnHeaders{iiHeader});
        end
    end
    if foundData
        tableOut = vertcat(tableOut, tablefromSheetOutClean);
    end
    % let's look for correct dates and exclude anything without a good date
    datetimeFormatString = 'dd-MMM-yyyy';
    tableOut.("DOB") = datetime(tableOut.("DOB"), 'InputFormat', datetimeFormatString);
    missingDates = isnat(tableOut.("DOB"));
    tableOut(missingDates,:) = [];
    warning(['File: ' xlsxFileName ' Sheet: ' sheetList{iSheet} ' excluded ' num2str(sum(missingDates)) ' records due to bad date formatting.']);
end


% many columns still have missing data. working on it.
