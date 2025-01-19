function [success] = design_and_k_calculate(input_macro_path)
    %clc;
close all;

addpath(genpath('C:\\Users\\.....\\app\\CST-MATLAB-API-master\\CST-MATLAB-API-master')); %this line should point to the folder where the CST-MATLAB-API-master repo is installed. For instance, my path is as given.
 

%This command is used to initiate the CST application, as you can see here, it is assigned to
%the cst variable. 
cst = actxserver('CSTStudio.application');

%This command is used to open a new CST project. From now on this mws
%variable will be used to operate any matlab function that is used for the
%control of CST

macroPath = fullfile('custom', input_macro_path);

mws = cst.invoke('NewMWS');
invoke(mws, 'RunMacro', macroPath)
invoke(mws, 'RunMacro', 'Solver\Ports\Calculate-port-extension-coefficient-MWS');
success = true;

end
