function [Total_DemandKWh,Total_NetKWh,Total_SolarKWh,Total_BatteryKWh,Demand_tRMDkW,Batt_tRMDkW,PV_tRMDkW, Net_tRMDkW, HMP, HMP_Batt, HMP_PV, tRMD, tHMP, checkRMD, checkHMP, dataOK] = EnergyvsNetEnergy_loop2(DATA,HourBegin,HourEnd,Daynum)
%EnergyvsNetEnergy_loop Calculates microgrid performance indicators for each day during
%peak hours.
%   Calculates the energy used, total renewable enrgy, etc.

% RMD = Reduction in Maximum Demand
%    At the time during peak hours when thedemand was the highest, what
%    was the achived % reduction?
% HMP = Highest Microgrid Performance
%    During peak hours, what was the largest achieved reduction in demand?

% Indicies of peak hours of day Daynum
select = DATA(:,6)>=HourBegin & DATA(:,6)<HourEnd & DATA(:,5)==Daynum ;
select2 = DATA(:,6)>=HourBegin + 0/60 & DATA(:,6)<HourEnd & DATA(:,5)==Daynum ; % Use this to exclude first datapoint of peak hours
sInds = find(select==1);
%sInds2 = find(select2==1);
% Indicies of select, for indexing during peak hours

%% Energy Calcs

Total_NetKWh=sum(DATA(select,7)); 
Total_DemandKWh=sum(DATA(select,8));
Total_BatteryKWh=sum(DATA(select,4));
Total_SolarKWh=sum(DATA(select,3));

%% Power Calcs

[Demand_tRMDkW,tRMD] = max(DATA(select2,8)); % tRMD is the time at which the demand was highest during peak hours

if length(tRMD)==1
    Demand_tRMDkW = Demand_tRMDkW*12;       % Highest demand during peak hours
    Net_tRMDkW = DATA(sInds(tRMD), 7)*12;           % Net demand (kW) at time when demand was highest ""
    Batt_tRMDkW = DATA(sInds(tRMD),4)*12;
    PV_tRMDkW = DATA(sInds(tRMD),3)*12;
    dataOK = true;
else
    % -999 so we can easily ID bad data later. 
    Batt_tRMDkW = -999;
    Net_tRMDkW = -999;
    PV_tRMDkW = -999;
    Demand_tRMDkW = -999;
    dataOK = false;
end

%(Demand_tRMDkW - Net_tRMDkW)/Demand_tRMDkW*100 % RMD

%% HMP
[HMP, tHMP] = max(DATA(select2, 8) - DATA(select2, 7));
if dataOK
    HMP = HMP*12;
    HMP = (HMP/(DATA(sInds(tHMP), 8)*12))*100;
    HMP_Batt = DATA(sInds(tHMP), 4)/DATA(sInds(tHMP),8)*100; % Battery contribution to HMP, %
    HMP_PV = DATA(sInds(tHMP), 3)/DATA(sInds(tHMP), 8)*100;  % PV contribution to HMP, %
    
    %HBP = Highest battery performance is differnt from HMP in that it
    %finds the time when the battery made the largest reduction in demand
    %HPVP = ^^ for PV
   
else
    HMP = -999;
    HMP_Batt = -999;
    HMP_PV = -999;
end


%% Check for Spikes
% tRMD = select(sInds(tRMD));
% tHMP = sInds(tHMP);
if dataOK
    threshold = 5;  % Allowable difference in battery power around tRMD/tHMP
    
    checkRMD = 0;
    checkHMP = 0;
    battery = DATA(sInds, 4).*12; % kW
    if or(tRMD == 1, tRMD == length(sInds))
        checkRMD = 1;
    else
        %if or(abs(battery(tRMD)-1)<3, abs(battery(tRMD+1)<3))
        if abs(battery(tRMD-1) - battery(tRMD+1)) > threshold
            checkRMD = 1;
        end
    end
    
    
    
    if or(tHMP == 1, tHMP == length(sInds))
        checkHMP = 1;
    else
        if abs(battery(tHMP-1) - battery(tHMP+1)) > threshold
            checkHMP = 1;
        end
    end
    
    % Output the excel time for tRMD and tHMP
    tRMD = DATA(sInds(tRMD), 1); 
    tHMP = DATA(sInds(tHMP), 1);
else % if data not ok
    checkRMD = 0;
    checkHMP = 0;
    tHMP = -999;
    tRMD = -999;
end


       

