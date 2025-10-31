function [tableOut,badRecordTable] = readAndCombineXlsxRecord(xlsxFileName,keyColumnHeaders,primaryHeader)
% given: a xls table path, header names, and a save file location
% DO: step through each sheet, pull out the column based on first header
% location, and combine them into a master table and return that table

% % example input
% % xlsxFileName = 'Z:\PearceLabRecords\Mouse Inventory\Lamp5-cre\Lamp5-cre.xlsx'; % we'll step through a list of these, but start here for testing
% keyColumnHeaders = {'ID Number','DOB','Date of Exp','mouseAssignment','sacCode','fundingID'};
% primaryHeader = 'DOB'; % in case the formatting is incomplete, we need one header that will determine if a record is valid or not.
% % problem child:
% xlsxFileName = 'Z:\PearceLabRecords\Mouse Inventory\2025 GABRb2\2025 GABRb2.xlsx'

% toggle verbose mode
showBadRecords = true;
badRecordTable = table;

tableOut = table('Size',[0,length(keyColumnHeaders)],'VariableTypes',{'string','string','string','string','string','string'},'VariableNames',keyColumnHeaders);
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
    % We will now search through the document for the best fit for the keycolumn headers
    keyColumnHeadersChecklist = zeros(size(rawSheetAsCell));
    for iRow = 1:size(rawSheetAsCell, 1)
        rowData = string(rawSheetAsCell(iRow, :));
        % if we find our primary header, look for the other headers
        if any(rowData == primaryHeader)
            for iHeader = 1:size(keyColumnHeaders, 2)
                if any(rowData == keyColumnHeaders{iHeader})
                    keyColumnHeadersChecklist(iRow,iHeader) = 1;
                end
            end
        end
    end
    [~,headerRow] = max(sum(keyColumnHeadersChecklist,2));
    rawSheetAsCell = rawSheetAsCell(headerRow:end,:);
    headerNamesTemp = string(rawSheetAsCell(1,:));
    theseEmpty = cellfun(@isempty, headerNamesTemp);
    uniquePlaceholders = compose('X%d', 1:sum(theseEmpty));
    headerNamesTemp(theseEmpty) = uniquePlaceholders;
    thisTable = cell2table(rawSheetAsCell(2:end,:), 'VariableNames', headerNamesTemp);

    % The old method would scramble columns if data were incomplete.  Let's add a "primary" header option, and just look for
    % rows that have a valid key header. . . . getting wonky, but so are
    % the data.  pruning based on DOB here.  This is no longer a flexible
    % function and instead depends on "DOB"
    datetimeFormatString = 'dd-MMM-yyyy';
    thisTable.(primaryHeader) = string(thisTable.(primaryHeader));
    invalidDOB = isnat(datetime(thisTable.(primaryHeader), 'InputFormat', datetimeFormatString));
    thisTable(invalidDOB,:) = [];
    tempTable = table('Size',[height(thisTable),length(keyColumnHeaders)],'VariableTypes',{'string','string','string','string','string','string'},'VariableNames',keyColumnHeaders);
    for iHeader = 1:size(keyColumnHeaders,2)
        if ismember(keyColumnHeaders{iHeader},thisTable.Properties.VariableNames)
            tempTable.(keyColumnHeaders{iHeader}) = string(thisTable.(keyColumnHeaders{iHeader}));
        end
    end

    tableOut = vertcat(tableOut, tempTable);
    
end






