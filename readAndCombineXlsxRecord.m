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

tableOut = table('Size',[0,length(keyColumnHeaders)],'VariableTypes',{'string','string','string','string','string','string'},'VariableNames',keyColumnHeaders);;
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



    % we're starting over... We will now search through the document for
    % the best fit for the keycolumn headers
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



    % for iCol = 1:size(rawSheetAsCell, 2)
    %     columnData = string(rawSheetAsCell(:, iCol));
    %     for iHeader = 1:size(keyColumnHeaders, 2)
    %         targetHeader = string(keyColumnHeaders{iHeader});
    %         singleColumn = (columnData == targetHeader);
    %         if any(singleColumn) 
    %             foundRow = find(singleColumn, 1, 'first');
    %             earliestHeaderRow = min(earliestHeaderRow, foundRow);
    %         end
    %     end
    % end
    % opts.DataLines = [earliestHeaderRow + 1, inf];
    % colNumber = size(rawSheetAsCell,2);
    % remainder = mod(colNumber - 1, 26);
    % % Convert the remainder (0-25) to an ASCII character (A=65)
    % colLetter = char('A' + remainder);
    % opts.DataRange = ['A', num2str(headerRow), ':', colLetter, num2str(size(columnData,1))]; 
    % 
    % % here, we'll fix all imported rows as 'char' so we can search it easily.
    % numColumns = size(rawSheetAsCell,2);
    % fixedType = {'char'};
    % charArray = repmat(fixedType, 1, numColumns);
    % opts.VariableTypes(1:numColumns) = charArray;
    % thisTable = readtable(xlsxFileName,opts); 

    % TODO: figure out how to load it once only
    % fixedType = {'char'};
    % numColumns = size(thisTable,2);
    % charArray = repmat(fixedType, 1, numColumns);
    % opts.VariableTypes(1:numColumns) = charArray;
    % opts.VariableNamingRule = 'preserve'; 

    % we could use opts.VariableNames to confirm we got the right columns?


    % Now search through and create a sub table of specified var from this sheet
    % thisTable = readtable(xlsxFileName,opts);

    % trying something new.  The old method would scramble columns if data
    % were incomplete.  Let's add a "primary" header option, and just look for
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




    % 
    % tablefromSheetOut = table('Size',[size(columnData,1),length(keyColumnHeaders)],'VariableTypes',{'string','string','string','string','string','string'},'VariableNames',keyColumnHeaders);
    % for iHeader = 1:size(keyColumnHeaders,2)
    %     % step through each column and look for matching key words until we
    %     % find what we're looking for (or run out of columns)
    %     foundData = false;
    %     tempTable = table;
    %     iCol = 1;
    %     while foundData == false 
    %         singleColumn = ismember(thisTable{:,iCol},keyColumnHeaders{iHeader});
    %         if sum(singleColumn)>0
    %             thisRow = find(singleColumn,1,'first');
    %             tempTable = thisTable(thisRow+1:end,iCol);
    %             tempTable = renamevars(tempTable, tempTable.Properties.VariableNames, keyColumnHeaders(iHeader));
    %             foundData = true;
    %             doOnce = true;
    %         end
    %         if iCol >= size(thisTable,2)
    %             foundData = true;
    %         end
    %         iCol = iCol+1;
    %     end
    %     if foundData & ~isempty(tempTable)
    %         % tablefromSheetOut = horzcat(tablefromSheetOut, tempTable);
    %         tablefromSheetOut.(keyColumnHeaders{iHeader})(1:size(tempTable.(keyColumnHeaders{iHeader}),1)) = tempTable.(keyColumnHeaders{iHeader});
    %     end
    % end
    % 
    % 
    % tablefromSheetOutClean = table('Size',[height(tablefromSheetOut),length(keyColumnHeaders)],'VariableTypes',{'string','string','string','string','string','string'},'VariableNames',keyColumnHeaders);
    % for iiHeader = 1:length(keyColumnHeaders)
    %     if ismember(keyColumnHeaders{iiHeader}, tablefromSheetOut.Properties.VariableNames)
    %         tablefromSheetOutClean.(keyColumnHeaders{iiHeader}) = tablefromSheetOut.(keyColumnHeaders{iiHeader});
    %     end
    % end

    


    tableOut = vertcat(tableOut, tempTable);



    % if foundData
    %     % let's look for correct dates and exclude anything without a good date
    %     datetimeFormatString = 'dd-MMM-yyyy';
    %     tablefromSheetOutClean.("DOB") = datetime(tablefromSheetOutClean.("DOB"), 'InputFormat', datetimeFormatString);
    %     missingDates = isnat(tablefromSheetOutClean.("DOB"));
    %     badRecordTable = vertcat(tablefromSheetOutClean(missingDates,:), badRecordTable);
    %     if showBadRecords
    %         disp(['File: ' xlsxFileName ' Sheet: ' sheetList{iSheet} ' excluded ' num2str(sum(missingDates)) ' records due to bad date formatting.']);
    %     end
    %     tablefromSheetOutClean(missingDates,:) = [];
    %     tableOut = vertcat(tableOut, tablefromSheetOutClean);
    % end


    

end


% many columns still have missing data. working on it.



% TODO!!  noticing a problem where DOB is entered below and some records
% are correctly included and others are not.  THIS NEEDS TO BE SOLVED but I
% ran out of time today 10/10/25.  I think we can just check for both ID
% and DOB for any row, and include it if they both exist?  try that and see
% if that helps.  and no need to include totally empty rows in the excluded
% records (they aren't even records, just bad input).
