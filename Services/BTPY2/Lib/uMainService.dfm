object SvcSyncBTPY2: TSvcSyncBTPY2
  OldCreateOrder = False
  DisplayName = 'Synchronisation BTP Y2'
  ServiceStartName = 'j.trifilieff@spare.local'
  AfterInstall = ServiceAfterInstall
  OnExecute = ServiceExecute
  OnStart = ServiceStart
  OnStop = ServiceStop
  Left = 198
  Top = 117
  Height = 150
  Width = 215
end
