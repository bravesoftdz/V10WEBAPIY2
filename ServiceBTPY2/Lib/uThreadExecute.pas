unit uThreadExecute;

interface

uses
  Classes;

type
  SynchroThread = class(TThread)
  private
    { Déclarations privées }
  protected
    procedure Execute; override;
  end;

  R_Params = record
    ServerName  : string;
    DBName      : string;
    LastSynchro : string;
  end;


implementation

{ Important : les méthodes et propriétés des objets de la VCL peuvent uniquement
  être utilisés dans une méthode appelée en utilisant Synchronize, comme :

      Synchronize(UpdateCaption);

  où UpdateCaption serait de la forme 

    procedure SynchroThread.UpdateCaption;
    begin
      Form1.Caption := 'Mis à jour dans un thread';
    end; }

{ SynchroThread }

procedure SynchroThread.Execute;
begin
  
  { Placez le code du thread ici }
end;

end.
