unit uThreadExecute;

interface

uses
  Classes;

type
  SynchroThread = class(TThread)
  private
    { D�clarations priv�es }
  protected
    procedure Execute; override;
  end;

  R_Params = record
    ServerName  : string;
    DBName      : string;
    LastSynchro : string;
  end;


implementation

{ Important : les m�thodes et propri�t�s des objets de la VCL peuvent uniquement
  �tre utilis�s dans une m�thode appel�e en utilisant Synchronize, comme :

      Synchronize(UpdateCaption);

  o� UpdateCaption serait de la forme 

    procedure SynchroThread.UpdateCaption;
    begin
      Form1.Caption := 'Mis � jour dans un thread';
    end; }

{ SynchroThread }

procedure SynchroThread.Execute;
begin
  
  { Placez le code du thread ici }
end;

end.
