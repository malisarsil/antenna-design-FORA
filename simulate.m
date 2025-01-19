function [success] = simulate(input_macro_path)
    %clc;
    close all;
    %clear all;
    
    addpath(genpath('C:\\Users\\.....\\CST-MATLAB-API-master\\CST-MATLAB-API-master')); %this line should point to the folder where the CST-MATLAB-API-master repo is installed. For instance, my path is as given.
    
    %This command is used to initiate the CST application, as you can see here, it is assigned to
    %the cst variable. 
    cst = actxserver('CSTStudio.application');
  
    mws = cst.invoke('NewMWS');
    macroPath = fullfile('custom', input_macro_path);
    invoke(mws, 'RunMacro', macroPath)
    
    CstDefineTimedomainSolver(mws,-40)
    
    CstDefineEfieldMonitor(mws,strcat('e-field', '2.45'),2.45);
    CstDefineHfieldMonitor(mws,strcat('h-field', '2.45'), 2.45);
    CstDefineFarfieldMonitor(mws,strcat('Farfield','2.45'), 2.45);
    
    
    exportpath = 'C:\\Users\\.....\\app\\result'; % Adjust to your directory
    filenameTXT = 'results.txt';

    % Export S11 as a .txt file
    CstExportSparametersTXT(mws, exportpath, filenameTXT);
    
    [Frequency, Sparametter] = CstLoadSparametterTXT(exportpath, filenameTXT);
    
    % Plot S-parameter
    figure;
    plot(Frequency, Sparametter);
    xlabel('Frequency (GHz)');
    ylabel('|S11| (dB)');
    title('S11 Parameter');
    grid on;
    CstSaveProject(mws)
    
    CstQuitProject(mws)
    success = true;
end
