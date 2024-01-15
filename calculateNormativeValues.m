%Calculate Normative Values
%- Calculates normative values (means, std, 95% CI) for each variable for each subject
%Saves entire table array as dataTable and individual sports tables as the
%sport name in normativeValues folder
clear
close all

%Load file with raw data for all speeds
[file, path] = uigetfile('*','Select the master data file to open'); %referenceData_osimOutput
data = importdata([path,filesep,file]); %import master file
rawData = data.data;
headers = data.textdata(1,:);
rawDataCell = num2cell(rawData);

%Sport List Generation
subsport_numect_num = rawData(:,1); %subsport_numect numbers from master file
subsport_numect_char = num2str(subsport_numect_num); %subsport_numect numbers from master file in strings
sport_char = subsport_numect_char(:,1:3); %first three numbers from each cell to determine sport/gender
sport_num = str2num(sport_char); %#ok<*ST2NM> %change index strings to numbers

%Add sport name to cell array
size(sport_num);
[dataRows, ~] = size(sport_num);
sportList = cell(dataRows, 1); %creates empty cell array
sportList(sport_num == 122) = {'M_Football'}; %add each sport name to empty cell array based on subject number
sportList(sport_num == 102) = {'M_Basketball'};
sportList(sport_num == 202) = {'M_CC'};
sportList(sport_num == 132) = {'M_Golf'};
sportList(sport_num == 162) = {'M_Soccer'};
sportList(sport_num == 232) = {'M_Tennis'};
sportList(sport_num == 192) = {'M_Track'};
sportList(sport_num == 222) = {'Wrestling'};
sportList(sport_num == 111) = {'Spirit'};
sportList(sport_num == 101) = {'W_Basketball'};
sportList(sport_num == 201) = {'W_CC'};
sportList(sport_num == 151) = {'W_Crew'};
sportList(sport_num == 141) = {'W_Hockey'};
sportList(sport_num == 161) = {'W_Soccer'};
sportList(sport_num == 171) = {'W_Softball'};
sportList(sport_num == 181) = {'W_Swim'};
sportList(sport_num == 191) = {'W_Track'};
sportList(sport_num == 231) = {'W_Tennis'};
sportList(sport_num == 211) = {'W_Volleyball'};

%Assigning footstrikes based on foot angle
footstkLeft_Column = find(strcmp(headers, 'footAngleContactL'));
footstkRight_Column = find(strcmp(headers, 'footAngleContactR'));
Lstrike_data = rawData(:,footstkLeft_Column);
Rstrike_data = rawData(:,footstkRight_Column);
strikeList(Lstrike_data <0) = {'Forefoot'}; %assigning footstrike pattern based on foot contact angle for each foot
strikeList(Lstrike_data >0) = {'Rearfoot'};
strikeList(Rstrike_data <0) = {'Forefoot'};
strikeList(Rstrike_data >0) = {'Rearfoot'};
strikeList = (strikeList)'; %List of all subject's footstrikes

%Adding sport and footstrike categories to master file
rawDataCell(:,end+1) = sportList; %adding sport list to end of master file
doublesportList = repmat(sportList,2); %repeat sport list in new column
rawDataCell(:,end+1) = strikeList;%adding footstrike list to end of master file
doublesktList = repmat(strikeList,2);%repeat footstrike list in new column

%Making arrays of only left and right foot data based on header names
[row, col] = size(rawDataCell);
leftRows = 'L';
leftVars = endsWith(headers, leftRows); %finding which left foot variable headers
leftData = rawDataCell(:,leftVars); %making array with only the left foot data
rightRows = 'R';
rightVars = endsWith(headers, rightRows); %finding which right foot variable headers
rightData = rawDataCell(:,rightVars); %making array with only the right foot data

%Duplicate demographic data for converting data from wide to long format
demographicData = rawDataCell(:,1:8);
doubledemoData = repmat(demographicData,2);
headers(:, leftVars) = [];
demographicHeaders = headers(:, 1:8);
demoHeadersSize = size(demographicHeaders);
newHeaders = strip(headers,'right','R');
newHeaders(:,end+1) = {'sport'};
newHeaders(:,end+1) = {'footstrike'};
newHeadersSize = size(newHeaders);

%Moving all data to new array
[dataRows, dataCols] = size(newHeaders);
dataLong = cell((row*2)+1, newHeadersSize(:,2));
dataLong(1,1:dataCols) = newHeaders;
dataLong(2:end, 1:demoHeadersSize(:,2)) = doubledemoData(:, 1:demoHeadersSize(:,2));
dataLong(2:(row+1), (demoHeadersSize(:,2))+1:end-2) = leftData;
dataLong(row+2:end, (demoHeadersSize(:,2))+1:end-2) = rightData;
dataLong(2:end, end-1) = doublesportList(:,1);
dataLong(2:end, end) = doublesktList(:,1);

%Normative value headers
normHeaders = {'Sport', 'Speed', 'Footstrike','Normative Data','Count','cadence','strides', 'stanceTime', 'stancePercent', 'strideLength', 'propImpulse', 'brakeImpulse',...
    'verticalImpulse', 'activePeak', 'impactPeak', 'loadingRate', 'mlForceNegPk', 'mlForcePosPk', 'heelComDistContact', 'baseGaitMidstance', 'vertComExcurTotal', 'vertComExcur',...
    'hipFlexContact', 'hipFlexPeak', 'hipFlexSwingPeak', 'hipExtPeak', 'hipExtPeakSwing', 'kneeFlexContact', 'kneeFlexPeak', 'kneeFlexSwingPeak', 'kneeExtSwingPeak', 'ankleAngleContact',...
    'ankleAngleToeOff', 'ankleDorsiPeak', 'footAngleContact', 'hipAddContact', 'hipAddPeak', 'pelvicTiltPeak', 'pelvicDropContact', 'pelvicDropPeak', 'pelvicRotSame','pelvicRotOpp',...
    'hipFlexMomPkSwing', 'hipFlexMomImpSwing', 'hipExtMomPk', 'hipExtMomImp', 'hipExtMomPkSwing', 'hipExtMomImpSwing','hipAbdMomPk', 'hipAbdMomImp', 'kneeFlexMomPeak', 'kneeFlexMomImp',...
    'kneeExtMomPeak', 'kneeExtMomImp', 'kneeExtMomRate2080', 'ankPFMomPeak', 'ankPFMomImp', 'hipNegWorkH0', 'kneeNegWorkK1', 'kneeNegWorkK4', 'ankNegWork', 'hipPosWorkH3','hipPosWorkH1',...
    'kneePosWorkK2', 'ankPosWork', 'vertStiffness', 'hipStiffness', 'kneeStiffness', 'ankleStiffness'};

%% Finding column numbers for sport, speed and footstrike.
% Cadence and ankle stiffness columns are needed for normative calculations
rawData = str2double(rawDataCell);
sportCol = find(strcmp(newHeaders, 'sport'));
sportData = dataLong(2:end,sportCol);
footstrikeCol = find(strcmp(newHeaders, 'footstrike'));
speedCol = find(strcmp(newHeaders, 'speedMs'));
cadenceColumn = find(strcmp(newHeaders, 'cadence'));
ankStiffColumn = find(strcmp(newHeaders, 'ankleStiffness'));

% Create a list of teams, speeds and footstrikes
teamList = {'M_Football', 'M_Basketball', 'M_CC', 'M_Golf','M_Soccer', 'M_Tennis', 'M_Track', 'Wrestling','Spirit', 'W_Basketball', 'W_CC', 'W_Crew', 'W_Hockey', 'W_Soccer', 'W_Softball','W_Swim', 'W_Track', 'W_Tennis','W_Volleyball'};
speedList = [268; 295; 335; 380; 412; 447];
footStrikeList = {'Forefoot', 'Rearfoot'};

% Create dataTable framework
[~, colValue] = size(normHeaders);
rowValue = length(teamList)*length(speedList)*length(footStrikeList)*4;
dataTable = cell(rowValue+1, colValue);
dataTable(1,:) = normHeaders;

%Format table with normative data row headers 
normLabels = repmat({'Mean', 'SD', 'Lower 95%', 'Upper 95%'}', rowValue/4, 1);
valueCol = find(strcmp(normHeaders, 'Normative Data'));
dataTable(2:end, valueCol) = normLabels;

%Create infoTable framework - gives number of subjects from each team
infoHeaders = {'Date', 'Sport','Speed', 'Footstrike','Number of Subjects'};
[~, Infocol] = size(infoHeaders);
infoRow = length(teamList)*length(speedList)*length(footStrikeList);
infoTable = cell(infoRow+1, Infocol);
infoTable(1,:) = infoHeaders;

%Normative data calculations and adding to sportTable
for t = 1:length(teamList)
    % Create sportTable framework
    [~, colValue] = size(normHeaders);
    rowValue = length(speedList)*length(footStrikeList)*4;
    sportTable = cell(rowValue+1, colValue);
    sportTable(1,:) = normHeaders;

    %Format table with normative data
    normLabels = repmat({'Mean', 'SD', 'Lower 95%', 'Upper 95%'}', rowValue/4, 1);
    sportTable(2:end, valueCol) = normLabels;

    sportRows = find(strcmp(sportData, teamList(t)) == 1);
    sportSubset = dataLong(sportRows+1,:);
    speedData = cell2mat(sportSubset(:,speedCol));
    for s = 1:length(speedList)
        speedRows = find(speedData == speedList(s));
        speedSubset = sportSubset(speedRows,:);
        fsData = cell2mat(speedSubset(:,footstrikeCol));
        for f = 1:length(footStrikeList)
            fsRows = find(strcmp(fsData, footStrikeList(f)) == 1);
            fsSubset = speedSubset(fsRows,:);
            countSubjects = height(fsSubset);

            sportStruct = teamList(t);
            speedStruct = append('S',num2str(speedList(s)));
            fsStruct = footStrikeList(f);

            blankData = cellfun(@isempty, infoTable);
            blankRows = find(blankData(:,2)==1);
            addRow = blankRows(1);

            %Add data to InfoTable
            infoTable(2,find(strcmp(infoHeaders, 'Date'))) = {date};
            infoTable(addRow, find(strcmp(infoHeaders, 'Sport'))) = {teamList(t)}; 
            infoTable(addRow, find(strcmp(infoHeaders, 'Speed'))) = {speedList(s)};
            infoTable(addRow, find(strcmp(infoHeaders, 'Footstrike'))) = {footStrikeList(f)};

            if countSubjects == 0
                infoTable(addRow, find(strcmp(infoHeaders, 'Number of Subjects'))) = {0};
            else
                infoTable(addRow, find(strcmp(infoHeaders, 'Number of Subjects'))) = {countSubjects};
            end

            emptyData = cellfun(@isempty, sportTable);
            emptySport = find(emptyData(:,1)==1);
            nextRow = emptySport(1);

            if isempty(fsSubset) == 1
                avgValue = zeros(length(cadenceColumn:ankStiffColumn));
                sdValue = zeros(length(cadenceColumn:ankStiffColumn));
                CIValueUp = zeros(length(cadenceColumn:ankStiffColumn));
                CIValueLow = zeros(length(cadenceColumn:ankStiffColumn));
            else
                for c = cadenceColumn:ankStiffColumn
                    strides = find(strcmp(normHeaders, 'strides'));

                    %Calculate mean, stdev, 95% confidence intervals
                    avgValue = mean(cell2mat(fsSubset(:,c)));
                    sdValue = std(cell2mat(fsSubset(:,c)));
                    fsRowsSize = size(fsRows);
                    CIValueUp = avgValue + ((sdValue)/sqrt(fsRowsSize(:,1)));
                    CIValueLow = avgValue - ((sdValue)/sqrt(fsRowsSize(:,1)));
                    
                    %Add footstike, speed, sport to table
                    sportTable(nextRow:nextRow+3, find(strcmp(normHeaders, 'Sport'))) = {teamList(t)}; %#ok<*FNDSB> 
                    sportTable(nextRow:nextRow+3, find(strcmp(normHeaders, 'Speed'))) = {speedList(s)};
                    sportTable(nextRow:nextRow+3, find(strcmp(normHeaders, 'Footstrike'))) = {footStrikeList(f)};
                    sportTable(nextRow:nextRow+3, find(strcmp(normHeaders, 'Count'))) = {countSubjects};
        
                    %Add normative values
                    sportTable(nextRow, c-2) = {avgValue};
                    sportTable(nextRow + 1, c-2) = {sdValue};
                    sportTable(nextRow + 2, c-2) = {CIValueLow};
                    sportTable(nextRow + 3, c-2) = {CIValueUp};
       
                end
            end

            %Storing data in a structure so can be used for plotting in
            %Data_Comparison
            [plotR, plotC] = size(fsSubset);
            plotDataTable = cell(plotR+1, plotC);
            plotDataTable(1,:) = newHeaders;
            plotDataTable(2:end,:) = fsSubset;
            plotting.(cell2mat(sportStruct)).((speedStruct)).(cell2mat(fsStruct)) = plotDataTable; %Not separating rear and fore
        end
    end
    %Save plotting data
    filepath = 'S:\Analysis\Gait Reporting\normativeValues\calculateNormativeValues\';
    plotFile = fullfile(filepath,'PlottingData.mat');
    save(plotFile, '-struct', 'plotting');

    %Save sport specific data
    filename = fullfile(filepath,'\Sport Specific Normative Data\',[char(teamList(t)), '.mat']);
    excelFile = [char(teamList(t)), '.xlsx'];
    save(filename, "sportTable")
    sportTable = cell2table(sportTable);
    sportPath = fullfile(filepath,'\Sport Specific Normative Data\',excelFile);
    writetable(sportTable, sportPath,'FileType','spreadsheet','Sheet', 'Normative Data');

    %Save infoTable 
    infoFile = fullfile(filepath,'infoTable.mat'); 
    save(infoFile, 'infoTable')
    infoTable = cell2table(infoTable);
    infoPath = fullfile(filepath,'infoTable.xlsx');
    writetable(infoTable, infoPath,'FileType','spreadsheet','Sheet', 'Info');
    infoTable  = table2cell(infoTable);

    clear nextRow

    %% Write sportTable to bigger dataTable
    % Define row dynamically
    emptyData = cellfun(@isempty,dataTable);
    emptySport = find(emptyData(:,1)==1);
    nextRow = emptySport(1);
    dataLength = size(sportTable,1)-1; % "-1" to account for first rows of headers

    sportTable = table2cell(sportTable);
    dataTable(nextRow:nextRow+dataLength-1, :) = sportTable(2:end,:);

    sportTable = cell2table(sportTable);
    normDataPath = fullfile(filepath,'normData.xlsx');
    writetable(sportTable,normDataPath,'FileType','spreadsheet','Sheet',char(teamList(t)));

    clear sportTable
end
dataTable = dataTable(~any(cellfun('isempty', dataTable), 2), :);

% This is outside all of the loops so it only saves dataTable once, after
% all data is written to it. 
normFile = 'normData.mat'; 
save(normFile, 'dataTable')
dataTable = cell2table(dataTable);
filepath = 'S:\Analysis\Gait Reporting\normativeValues\calculateNormativeValues\';
dataTablePath = fullfile(filepath,'normData.xlsx');
writetable(dataTable,dataTablePath,'FileType','spreadsheet','Sheet', 'All Data'); %All data

%Deleting extra char(teamList(t)) sheet
excelFileName = 'normData.xls';
excelFilePath = pwd; % Current working directory.
sheetName = 'char(teamList(t))'; % EN: Sheet, DE: Tabelle, etc. (Lang. dependent)
% Open Excel file.
objExcel = actxserver('Excel.Application');
objExcel.Workbooks.Open(fullfile(excelFilePath, excelFileName)); % Full path is necessary!
% Delete sheets.
try
      % Throws an error if the sheets do not exist.
      objExcel.ActiveWorkbook.Worksheets.Item(sheetName).Delete;
catch
      % Do nothing.
end
% Save, close and clean up.
objExcel.ActiveWorkbook.Save;
objExcel.ActiveWorkbook.Close;
objExcel.Quit;
objExcel.delete;

%% Creating data arrays based on gender
load("normData.mat") %Loads as dataTable
femaleSport = {'W_Basketball', 'W_CC', 'W_Crew', 'W_Hockey', 'W_Soccer', 'W_Softball','W_Swim', 'W_Track', 'W_Tennis','W_Volleyball'};
maleSport = {'M_Football', 'M_Basketball', 'M_CC', 'M_Golf','M_Soccer', 'M_Tennis', 'M_Track', 'Wrestling'};
teamCol = find(strcmp(dataTable(1,:), 'Sport'));
teamAll = dataTable(:, teamCol);
teamAll = vertcat(teamAll{:});
dataTableHeaders = dataTable(1,:);

femaleTeams = ismember(teamAll, femaleSport);
femaleData = [dataTableHeaders;dataTable(femaleTeams == 1,:)];
femaleData(1,:) = dataTableHeaders;
genderFile = fullfile(filepath,'femaleData.mat'); 
save(genderFile, 'femaleData');

maleTeams = ismember(teamAll, maleSport);
maleData = [dataTableHeaders;dataTable(maleTeams == 1,:)];
maleData(1,:) = dataTableHeaders;
genderFile = fullfile(filepath,'maleData.mat'); 
save(genderFile, 'maleData');