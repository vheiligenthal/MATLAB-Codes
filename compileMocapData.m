function compileMocapData

clc
clear

%Define where you want to save the master excel file (current folder)
%inpath = uigetdir('S:\','Select a save location for the compiled dataset.');
inpath = pwd;
d = datetime('today');
d = datestr(d);

% Select what data set you want to compile
dataset = menu('What do you create a dataset based on?','Injury','Sport');
if dataset == 1
   
    %Injury Loop
    GaitMaster3D = readtable('3D Gait Master.xlsx','Sheet','Collection Info');
    sportCol = find(string(GaitMaster3D.Properties.VariableNames) =='Sport');
    typeCol = find(string(GaitMaster3D.Properties.VariableNames) =='InjuryType');
    locationCol = find(string(GaitMaster3D.Properties.VariableNames) =='InjuryLocation');
    sexCol = find(string(GaitMaster3D.Properties.VariableNames) =='Gender');
    GaitMaster3D = table2cell(GaitMaster3D);
    injuryType = {'ACL', 'BSI', 'Knee', 'Hip', 'Syndesmosis'};
    injuryList = listdlg('PromptString','For which injury type are you compiling data?','SelectionMode','multiple','ListString',injuryType);
    
    %Pulls subjects based on injury selection to add to master
    if injuryList == 1
        injury = ismember(GaitMaster3D(:, typeCol), 'ACL Tear');
        subjPull = GaitMaster3D(injury,:);
    elseif injuryList == 2
        injury = ismember(GaitMaster3D(:, typeCol), 'Bone Stress Injury');
        subjPull = GaitMaster3D(injury,:);
    elseif injuryList == 3
        injury = ismember(GaitMaster3D(:,typeCol),{'Cartilage Lesion', 'Meniscal Lesion', 'Ligament Injury'}) & strcmp(GaitMaster3D(:,locationCol), 'Knee');
        subjPull = GaitMaster3D(injury,:);
    elseif injuryList == 4
        injury = ismember(GaitMaster3D(:,typeCol),{'Impingement'}) & strcmp(GaitMaster3D(:,locationCol), 'Hip');
        subjPull = GaitMaster3D(injury,:);
    elseif injuryList == 5
        injury = ismember(GaitMaster3D(:,typeCol),{'Tendon Injury', 'Ligament Injury'}) & strcmp(GaitMaster3D(:,locationCol), 'Ankle');
        subjPull = GaitMaster3D(injury,:);
    end


    athletes = cell2mat(subjPull(:,1)); % Subjects from gait master that will be used in injury-based dataset 
    collections = cell2mat(subjPull(:,2)); % Collections from gait master that will be used in injury-based dataset
    athleteSports = subjPull(:, sportCol); 
    athleteSports = cellstr(athleteSports);
    athleteSex = subjPull(:,sexCol);
    
    load('sportFolder.mat') %loads as sportFolder
    [rows, ~] = size(athletes); %create empty array
    projectPaths = cell(rows, 1);
    for j = 1:length(athletes)
        sexMatch = find(strcmp(athleteSex(j), sportFolder(:,1)) == 1);
        sportMatch = find(strcmp(athleteSports(j), sportFolder(:,2)) == 1);
        folderMatch = intersect(sexMatch, sportMatch);
        projectPaths(j,1) = append('S:\Data Collection\3D Motion Data\Athletics\', sportFolder(folderMatch,3),'\');%Creates array of all projectPaths for all subjects
    end

    projectPaths = char(projectPaths); 
    
else
    % Sport loop 
    % Select the subjects/athletes you want to compile into the master spreadsheet
    % Select the sport you want to compile into the master spreadsheet
    contents = dir('S:\Data Collection\3D Motion Data\Athletics\'); % Finds all folders within Athletics, so you do not need to update a sport list when new sports are added
    projectDir = {contents.name}; % Extract folder names
    projectFolders = projectDir(3:end); % Remove first 2 entries which are "." and ".."
    projectList = listdlg('PromptString','For which sports are you compiling data?','SelectionMode','multiple','ListString',projectFolders);
    projectLocation = (['S:\Data Collection\3D Motion Data\Athletics\',char(projectFolders(projectList)),'\']);

    fileFolder = cellstr(ls([projectLocation,'*00*']));
    [subjects,~] = listdlg('PromptString','Select the subjects to compile','SelectionMode','multiple','ListString',fileFolder);

    % Subject loop - IDs data for each subject to write to the master excel file
    [~, collectCols] = size(subjects); %create empty array
    collections = cell(collectCols, 1);
    athletes = cell(collectCols, 1);
    for i = 1:length(subjects)
        athletes(i) = {char(fileFolder(subjects(i)))}; %#ok<NOPRT>
        if strlength(athletes(i)) == 8
            collections(i) = {'1'};
        else
            collections(i) = {athletes{i}(10:end)};
        end
        %Erases collection portion
        athletes(i) =  erase(athletes(i), athletes{i}(9:end));
    end
    projectPaths = repmat(projectLocation, length(athletes),1);
    athletes = str2double(athletes);
    collections = str2double(collections);  
end

%% Select if compiling running or jump data
trial_options = {'Running','Jump Forces','Jump Kinetics'};
trial_choice = listdlg('PromptString','Would you like to compile running or jumping data?','SelectionMode','multiple','ListString',trial_options);

% Loops through athletes from either injury or sport loop to run through
% compile functions
for k = 1:length(athletes)
    athlete = athletes(k);
    collection = collections(k);
    if collection == 1
        session = append(num2str(athlete));
    else
        session = append(num2str(athlete),'_',num2str(collection));
    end
    
    projectPath = projectPaths(k,:);
    subjFolder = append(projectPath,session,'\'); %Check if file exists, if not need to process
    collectionParams = append(projectPath,session,'\',session,'_collectionParams.mat');
    val = exists(collectionParams, 'file');
    
    if val == 0
        disp('OpenSim has not been processed for', session)
        continue
    else
        subjStruct = load(collectionParams);
    
        height = subjStruct.height;
        mass = load([subjFolder,'MoCap\mass.txt']);
        mass = round(mass,1);
        
        %Converts subject variable into a single cell value
        subjectCell = {session};
        
        %Selects which combination of compiling data will be done
        if trial_choice == 1
            compile_running(athlete, subjFolder, subjectCell, projectPath, inpath, d, height, mass,collection);
        elseif trial_choice== 2
            compile_jumpForces(athlete, subjFolder, subjectCell, inpath, d, height, mass,collection);
        elseif trial_choice == 3
            compile_jumpKinetics(athlete,subjFolder, subjectCell, inpath, d, height, mass,collection);
        elseif (length(trial_choice) == 2) && (trial_choice(1) == 1) && (trial_choice(2) == 2)
            compile_running(athlete, subjFolder, subjectCell, projectPath, inpath, d, height, mass,collection);
            compile_jumpForces(athlete,subjFolder, subjectCell, inpath, d, height, mass,collection);
        elseif (length(trial_choice) == 2) && (trial_choice(1) ==1) && (trial_choice(2) == 3)
            compile_running(athlete, subjFolder, subjectCell, projectPath, inpath, d, height, mass,collection);
            compile_jumpKinetics(athlete,subjFolder, subjectCell, inpath, d, height, mass,collection);
        elseif (length(trial_choice) == 2) && (trial_choice(1) ==2) && (trial_choice(2) == 3)
             compile_jumpForces(athlete,subjFolder, subjectCell, inpath, d, height, mass,collection);
            compile_jumpKinetics(athlete,subjFolder, subjectCell, inpath, d, height, mass,collection);
        else
            compile_running(athlete, subjFolder, subjectCell, projectPath, inpath, d, height, mass,collection);
            compile_jumpForces(athlete,subjFolder, subjectCell, inpath, d, height, mass,collection);
            compile_jumpKinetics(athlete,subjFolder, subjectCell, inpath, d, height, mass,collection);
        end
    end
end