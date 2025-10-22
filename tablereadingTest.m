% **** outline of mouse table curating we need to perform ****
% 1. load in the "table of tables" so we know what to operate on
% 2. use readAndCombineXlsxRecord 
% 2a. for each sheet, load it in with appropriate opts
% 2b. find the canonical columns and rows of interest (they don't all start in the same location!)
% 2c. merge what we found from each sheet
% 3. step through each sheet and merge!
addpath('\\anesfs1\home$\smgrady\settings\Documents\Code\Pearce-Lab-Utilities');


keyColumnHeaders = {'ID Number','DOB','Date of Exp','mouseAssignment','sacCode','fundingID'};
bigSaveFile = 'Z:\PearceLabRecords\Mouse Inventory\2025totalMouseCount.xlsx';


% example single animal file save
% xlsxFileName = 'Z:\PearceLabRecords\Mouse Inventory\Lamp5-cre\Lamp5-cre.xlsx'; 
% singleAnimalSaveFileName = 'Z:\PearceLabRecords\Mouse Inventory\Lamp5-cre\mouseCount.xlsx';

tableOfTablesFileName = 'Z:\PearceLabRecords\Mouse Inventory\InventorySummaryTest.xlsx';
opts = detectImportOptions(tableOfTablesFileName,'Sheet','Lines and paths');
opts.VariableNamingRule = 'preserve'; 
tableOfTables = readtable(tableOfTablesFileName,opts);
% loop through our list of tables!
bigTable = table;
allBadRecords = table;
for i = 1:size(tableOfTables,1)
    xlsxFileName = tableOfTables.("Full path"){i};
    thisLine = tableOfTables.("shorthand"){i};
    try
        disp(['Reading in ' thisLine]);
        % our new function
        [singleAnimaltable,badRecordTable] = readAndCombineXlsxRecord(xlsxFileName,keyColumnHeaders);
        % writetable(singleAnimaltable,singleAnimalSaveFileName);
        bigTable = vertcat(bigTable,singleAnimaltable);
        allBadRecords = vertcat(allBadRecords,badRecordTable);
        disp(['successfully read in in ' thisLine]);
    catch
        disp(['failed to read ' thisLine]);
        keyboard;
    end
end

writetable(bigTable,bigSaveFile,'Sheet','MouseSummary');
writetable(allBadRecords,bigSaveFile,'Sheet','ExcludedEntries');
