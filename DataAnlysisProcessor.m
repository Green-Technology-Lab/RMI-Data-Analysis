clear
close all

% load the data file
data = xls2struct('2_28_2019-2_28_2020 Load2.xlsx', 'Sheet1');

% Save the output as...
outFilename = 'Analyzed Data 2020-11-16-4.xlsx';

timestep = 5;  % Minutes (data resolution)
timestep = 60/timestep; % Factor to convery from kW to kWh.


convertTime = 693960;  % Difference between matlab datenum and excel datenum
Time = datenum(data.Time_Stamp+convertTime); % Convert from Excel to Matlab days

%% Load data from struct into some variables
% Years=year(datetime(Time,'ConvertFrom','datenum'));
% Days=day(datetime(Time,'ConvertFrom','datenum'));
Hours=hour(datetime(Time,'ConvertFrom','datenum'));
% Months=month(datetime(Time,'ConvertFrom','datenum'));
Demand = data.Demand; %Building demand
Daynums=floor(Time);
NetkWh=data.Net_Demand/timestep;

% Optional data filtering
%demand2 = smoothdata(data.Demand, 'gaussian', 5);

% Convert from kW to kWh.
DemandkWh=demand2/timestep;
DemandkWhNoSolar=(Demand-data.Solar)/timestep;
BatterykWh=data.Battery/timestep;
SolarkWh=data.Solar/timestep;

% Create a new data array
DATA=[Time,data.Net_Demand,SolarkWh,BatterykWh,Daynums,Hours,NetkWh,DemandkWh];

count=1;  % Need a seperate index for looped variables because loop index doesn't start at 1.   

%% Calculate peak-time energy metrics for each day
for i=floor(Time(1)):floor(Time(end))
    
    [Total_DemandkWh(count,1),Total_NetkWh(count,1),Total_SolarkWh(count,1),Total_BatterykWh(count,1),Demand_tRMDkW(count,1),Batt_tRMDkW(count,1),PV_tRMDkW(count,1), Net_tRMDkW(count,1), HMP(count,1), HMP_Batt(count,1), HMP_PV(count,1), tRMD(count,1), tHMP(count,1), checkRMD(count,1), checkHMP(count,1), dataOK(count,1)]=peakCalcs(DATA,16,21,i);
    outDay(count,1) = i-convertTime;  % Convert back to Excel days
    count=count+1;
end

% Convert to excel time. 
tRMD = tRMD - convertTime;
tHMP = tHMP - convertTime;



%% Energy Metrics

% Also calculate the percent of battery that was charged from PV

Total_DERkWh = Total_BatterykWh+Total_SolarkWh;               % Total energy use met by DERs during peak hours
DERkWhPercent = (Total_DERkWh)./Total_DemandkWh*100;             % Percentage of energy use during peak hours met by renewables
BatteryPercent = Total_BatterykWh./Total_DemandkWh*100; % Percentage of peak-time energy use met by battery
SolarPercent = Total_SolarkWh./Total_DemandkWh*100;     % Percentage of peak-time energy use met by PV

% Stats on energy metrics

% PV + Battery, percent of energy usage during peak statistics 
DERkWh_Per_Max = max(DERkWhPercent(dataOK));    
DERkWh_Per_avg = mean(DERkWhPercent(dataOK));
DERkWh_Per_std = std(DERkWhPercent(dataOK));

% Battery percent of energy usage during peak statistics 
BattkWh_Per_max = max(BatteryPercent(dataOK));
BattkWh_Per_avg = mean(BatteryPercent(dataOK));
BattkWh_Per_std = std(BatteryPercent(dataOK));

% PV percent of energy usage during peak statistics 
PVkWh_Per_max = max(SolarPercent(dataOK));
PVkWh_Per_avg = mean(SolarPercent(dataOK));
PVkWh_Per_std = std(SolarPercent(dataOK));


%% RMD
RMD = (Demand_tRMDkW - Net_tRMDkW)./Demand_tRMDkW*100; % RMD
% Each DER contribution to RMD
RMD_Batt = Batt_tRMDkW./Demand_tRMDkW.*100;      % Percent reduction of Max peak demand due to battery, or Battery-only RMD
RMD_PV = PV_tRMDkW./Demand_tRMDkW*100; % Percent reduction of Max peak demand due to Pv, or PV-only RMD


% Do statistics on RMD and HMP
RMD_max = max(RMD(dataOK));
RMD_avg = mean(RMD(dataOK));
RMD_std = std(RMD(dataOK));

RMD_Batt_max = max(RMD_Batt(dataOK));
RMD_Batt_avg = mean(RMD_Batt(dataOK));
RMD_Batt_std = std(RMD_Batt(dataOK));

RMD_PV_max = max(RMD_PV(dataOK));
RMD_PV_avg = mean(RMD_PV(dataOK));
RMD_PV_std = std(RMD_PV(dataOK));



%% HMP
% Each DER contribution to HMP?
HMP_max = max(HMP(dataOK));
HMP_avg = mean(HMP(dataOK));
HMP_std = std(HMP(dataOK));

HMP_Batt_max = max(HMP_Batt(dataOK));
HMP_Batt_avg = mean(HMP_Batt(dataOK));
HMP_Batt_std = std(HMP_Batt(dataOK));

HMP_PV_max = max(HMP_PV);
HMP_PV_avg = mean(HMP_PV(dataOK));
HMP_PV_std = std(HMP_PV(dataOK));


%% Save Results to File
% We can't write headers to excel files with numerics in one go, so have to
% write it row by row. This is slow, sorry. 
% To improve code running speed, comment out this section and view data in
% matlab. 

% Generate the header row
headers = {'Date','Total_DemandkWh','Total_NetkWh','Total_SolarkWh',...
    'Total_BatterykWh', 'Total_DERkWh', '',	'DERkWhPercent', ...
    'BatteryPercent','SolarPercent', '','DERkWh_Per_Max', ...
    'DERkWh_Per_avg','DERkWh_Per_std','BattkWh_Per_max','BattkWh_Per_avg',...
    'BattkWh_Per_std','PVkWh_Per_max','PVkWh_Per_avg','PVkWh_Per_std', ....
    '','Demand_tRMDkW','Batt_tRMDkW','PV_tRMDkW','Net_tRMDkW','',...
    'RMD','RMD_Batt','RMD_PV','','RMD_max', 'RMD_avg','RMD_std','',...
    'RMD_Batt_max','RMD_Batt_avg','RMD_Batt_std','','RMD_PV_max', ...
    'RMD_PV_avg', 'RMD_PV_std', '','HMP','HMP_Batt', 'HMP_PV','HMP_max',...
    'HMP_avg','HMP_std','HMP_Batt_max','HMP_Batt_avg','HMP_Batt_std',...
    '',	'HMP_PV_max','HMP_PV_avg','HMP_PV_std', '', 'tRMD', 'tHMP', ...
    'checkRMD', 'checkHMP', 'dataOK'};
blanks = zeros(size(RMD));


% Create the sheet and write the headers
xlswrite(outFilename, headers,1)

% Loop through the data, writing one row at a time
for i=1:length(outDay)
    t=now;
    % List of cells in each row to write
    location = strcat('A', num2str(i+1),':', 'BI', num2str(i+1));
    
    % The data to be written in each row
    xlswrite(outFilename, [outDay(i),Total_DemandkWh(i),Total_NetkWh(i),...
        Total_SolarkWh(i),Total_BatterykWh(i),Total_DERkWh(i), blanks(i),...
        DERkWhPercent(i), BatteryPercent(i),SolarPercent(i),blanks(i),...
        DERkWh_Per_Max,DERkWh_Per_avg,DERkWh_Per_std,...
        BattkWh_Per_max,BattkWh_Per_avg,BattkWh_Per_std,...
        PVkWh_Per_max,PVkWh_Per_avg,PVkWh_Per_std,blanks(i),...
        Demand_tRMDkW(i),Batt_tRMDkW(i),PV_tRMDkW(i),Net_tRMDkW(i),blanks(i)...
        ,RMD(i),RMD_Batt(i),RMD_PV(i),blanks(i),RMD_max,RMD_avg,...
        RMD_std,blanks(i),RMD_Batt_max,RMD_Batt_avg,RMD_Batt_std...
        ,blanks(i),RMD_PV_max,RMD_PV_avg,RMD_PV_std,blanks(i),...
        HMP(i),HMP_Batt(i),HMP_PV(i),HMP_max,HMP_avg,HMP_std,...
        HMP_Batt_max,HMP_Batt_avg,HMP_Batt_std,blanks(i),...
        HMP_PV_max,HMP_PV_avg,HMP_PV_std,blanks(i), tRMD(i), tHMP(i),...
        checkRMD(i), checkHMP(i), dataOK(i)],1,location);
    
    
    dt(i) = now-t;
    progress = i/length(outDay);
    eta = (mean(dt)*(length(outDay)-i)*1E5)/60;
    str = strcat({'Writing file '}, num2str(progress*100), {'%.'},{' Time Remaining: '}, num2str(eta), {' minutes.'});
    disp(str);
    
% Optional code to display file writing progress.
%     progress = i/length(outDay);
%     str = strcat({'Writing file '}, num2str(progress*100), {'%.'});
%     disp(str)
    
end
