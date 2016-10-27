clear;
folder_name = fullfile(uigetdir('H:\DSP\Radio-oncologie\commun\Physique radio-onco\Curiethérapie\Workload'));
cd(folder_name);

ucvoFiles=dir('*.csv');
string_time='Total treatment time (sec.):';

[~, computer] = system('hostname');
[~, user] =dos('ECHO %USERNAME%');


%Find offset
%Data is usually in E or F
offset1='F';
offset0='E';

for i=1:length(ucvoFiles)
    ActivityLine=24;
    Unitline=12;
    Dateline=22;
    
    
    offset=0;
    name=ucvoFiles(i,1).name;
    ptID=xlsread(name,1,'E2');
    killexcel(computer,user);
    if isempty(ptID)
        ptID=xlsread(name,1,'F2');
        killexcel(computer,user);
        offset=1;
    end

if offset==1
    cellAct=strcat(offset1,num2str(ActivityLine));
    cellUnit=strcat(offset1,num2str(Unitline));
    cellDate=strcat(offset1,num2str(Dateline));
    
    
end

if offset==0
    cellAct=strcat(offset0,num2str(ActivityLine));
    cellUnit=strcat(offset0,num2str(Unitline));
    cellDate=strcat(offset0,num2str(Dateline));
    
end
[blah2,Datetxt]=xlsread(name,1,cellDate);     %Set Date
killexcel(computer,user);

%If plan was not approved before saving csv everything is shifted up by 1.
if isempty(Datetxt)
    %sss={'DATE ERROR'};
    %Date(i)=sss;
    ActivityLine=ActivityLine-1;
    Unitline=Unitline-1;
    Dateline=Dateline-1;
    
    
    
else
    Date(i)=Datetxt;
end

ID(i)=ptID;                           %Set ID
Activity(i)=xlsread(name,1,cellAct);  %Set activity
killexcel(computer,user);
[blah1 Unittxt]=xlsread(name,1,cellUnit);
killexcel(computer,user);
Unit(i)=Unittxt;                      %Set Unit

[blah2,Datetxt]=xlsread(name,1,cellDate);     %Set Date
killexcel(computer,user);

if isempty(Datetxt)
    sss={'DATE ERROR'};
    Date(i)=sss;
else
    Date(i)=Datetxt;
end
    
    for j=45:100
        cell=strcat('A',num2str(j));
        [num,text]=xlsread(name,1,cell);
        killexcel(computer,user);
        
        
        if strcmp(text,string_time)
            celldata=strcat('D',num2str(j));
            totaltime(i)=xlsread(name,1,celldata);
            killexcel(computer,user);
            break
        end
        
    end

    
    
    sprintf('%d',ptID)
    
end

IDcell=num2cell(ID');
Actcell=num2cell(Activity');
timecell=num2cell(totaltime');
X=[{'Patient ID','Treatment Date','HDR Unit','Activity (Ci)','Treatment Time (s)'};IDcell,Date',Unit',Actcell,timecell];

%[file,path] = uiputfile('*.mat','Save Workspace As')
%folder_output= fullfile(uigetdir('H:\DSP\Radio-oncologie\commun\Physique radio-onco\Curiethérapie\Workload'));
xlswrite('RESULTS.xls',X)