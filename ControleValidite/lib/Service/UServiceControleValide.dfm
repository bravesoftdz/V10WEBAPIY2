object LSEAutorise: TLSEAutorise
  OldCreateOrder = False
  DisplayName = 'Service autorisation L.S.E'
  Interactive = True
  AfterInstall = ServiceAfterInstall
  OnContinue = ServiceContinue
  OnExecute = ServiceExecute
  OnPause = ServicePause
  OnShutdown = ServiceShutdown
  OnStart = ServiceStart
  OnStop = ServiceStop
  Left = 482
  Top = 327
  Height = 150
  Width = 215
  object TT: TTimer
    Interval = 10000
    OnTimer = TTTimer
    Left = 24
    Top = 32
  end
end
