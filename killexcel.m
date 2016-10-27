function killexcell(computer,user)
% [~, computer] = system('hostname');
% [~, user] = system('whoami');
pause(0.2)
[~, alltask] = system(['tasklist /S ', computer, ' /U ', user]);
excelPID = regexp(alltask, 'EXCEL.EXE\s*(\d+)\s', 'tokens');
for i = 1 : length(excelPID)
      killPID = cell2mat(excelPID{i});
      system(['taskkill /f /pid ', killPID]);
end