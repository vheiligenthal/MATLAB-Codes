function compile_running(subject, ~, session, projectPath, inpath, d, height, mass, collection)

subject = num2str(subject);
session = cell2mat(session);
%projectPath = cell2mat(projectPath);
collection = num2str(collection);
filename = [inpath, '\compiled_data_', d, '.xlsx'];

val = exist(filename, 'file');
if val == 0
    [~, txt, ~] = xlsread('report_headers.xlsx', 'running');
    headers = txt;
    %Headers for left limb
    xlswrite(filename, headers(:, 1)', 'all_speeds', 'A1');
    %Headers for right limb
    xlswrite(filename, headers(:, 2)', 'all_speeds', 'BT1');
    %Write key tab
    [~, txt, ~] = xlsread('report_headers.xlsx', 'runningKey');
    keyTxt = txt;
    xlswrite(filename, keyTxt, 'runningKey', 'A1');
    duplicates = 'Check for duplicate subjects';
    duplicates = repmat( {duplicates}, 1, 1);
    xlswrite(filename, duplicates, 'READ ME', 'A1');
else
    [~, B] = xlsfinfo(filename);
    sheetValid = any(strcmp(B, 'all_speeds'));
    if sheetValid == 0
        [~, txt, ~] = xlsread('report_headers.xlsx', 'running');
        headers = txt;
        %Headers for left limb
        xlswrite(filename, headers(:, 1)', 'all_speeds', 'A1');
        %Headers for right limb
        xlswrite(filename, headers(2:end, 2)', 'all_speeds', 'BT1');
        %Write key tab
        [~, txt, ~] = xlsread('report_headers.xlsx', 'runningKey');
        keyTxt = txt;
        xlswrite(filename, keyTxt, 'runningKey', 'A1');
        duplicates = 'Check for duplicate subjects';
        duplicates = repmat( {duplicates}, 1, 1);
        xlswrite(filename, duplicates, 'READ ME', 'A1');
    end
end

subjFile = [projectPath,num2str(session),'\Data Results\',num2str(session),'_ResearchReport.xlsx'];      
val = exist(subjFile, 'file');

if val == 0
    disp(['Running data for ', num2str(session), ' has not been processed.'])  
else
    [~, subjTab] = xlsfinfo(subjFile);
    %Finds tabs to exlude from writing to the running variables master,
    %i.e. "Sheet_1" and "jumps"
    tabsSheet = strfind(subjTab,'Sheet');
    tabsJump = strfind(subjTab,'jump');


    %Combines "sheet" and "jump" tab arrays.
    tabsExclude = [tabsSheet; tabsJump];
    
    %Replaces blank cells with zeros.
    tabsExclude(cellfun(@isempty, tabsExclude)) = {0};
    
    %Adds arrays together into a single row. This is needed to determine
    %which tab names contain neither "sheet" nor "jump", i.e. running tabs
    tabsCondense = cell2mat(tabsExclude);
    [~, c] = size(tabsCondense);
    tabsReduced = zeros(1, c);
    for a = 1:c
        tabsReduced(a) = sum(tabsCondense(:, a));
    end
    
    %Goes back to subj_tab variable and IDs which tabs to use based on the
    %tabs_reduced variable. If a value for a given cell is a number, it is
    %not a running data tab and will be skipped over when writing running
    %data to master dataset.
    
    for a = 1:c
        if tabsReduced(a) < 1
            tabsInclude(a) = subjTab(a); %#ok<AGROW>
        end
    end
    
    %Removes empty cells (i.e. non-running tabs) from tabs_include
    emptyCells = cellfun('isempty', tabsInclude);
    tabsInclude(emptyCells) = [];
    
    for t = 1:length(tabsInclude)
        
        speed = char(tabsInclude(t));
        
        %Find next empty row in master excel file to which new data will be
        %written
        %   [x, y] = xlsread(filename, 'all_speeds');
        [x, ~] = xlsread(filename, 'all_speeds');
        nRowsX = (size(x, 1));
        emptyRow = num2str(nRowsX+2);
        emptyCell = strcat('A', emptyRow);
        
        
        %Load variables from subject's excel file
        [~, dateProc, ~] = xlsread(subjFile, char(tabsInclude(t)), 'B2');
        [demoTemp, ~, ~] = xlsread(subjFile, char(tabsInclude(t)), 'B4:B5');
        [leftData, ~, ~] = xlsread(subjFile, char(tabsInclude(t)), 'C7:C69');
        [rightData, ~, ~] = xlsread(subjFile, char(tabsInclude(t)), 'D7:D69');
        
        %Transpose data from rows to columns
        leftData = num2cell(leftData)';
        rightData = num2cell(rightData)';
        
        %% Concatenate data
        %Reduces number of times xlswrite is called & speeds up the code    
        emptyCellLeft = strcat('I', emptyRow);
        emptyCellRight = strcat('BT', emptyRow);
        
        demo_info = {subject(1:8), collection, height, mass, char(dateProc), speed, demoTemp(1), demoTemp(2)}; % separates ID number and collection number
        
        xlswrite(filename, demo_info, 'all_speeds', emptyCell);
        xlswrite(filename, leftData, 'all_speeds', emptyCellLeft);
        xlswrite(filename, rightData, 'all_speeds', emptyCellRight);
    end  
end
clear tabsInclude