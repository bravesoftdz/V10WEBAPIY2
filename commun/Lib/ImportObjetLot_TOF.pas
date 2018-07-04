{***********UNITE*************************************************
Auteur  ...... :  JS
Créé le ...... : 30/09/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : YYIMPORTOBJETLOT ()
                 Import de bob multi-sociétés
Mots clefs ... : TOF;IMPORTOBJETLOT
*****************************************************************}
Unit ImportObjetLot_TOF ;

Interface

Uses StdCtrls,
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF} 
     mul, MajTable,
{$else}
     eMul,
     uTob, UtileAGL,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1, MenuOlg,
     HMsgBox, HSysMenu,
     UTOF,
      UBob ;

Type
  TOF_IMPORTOBJETLOT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

    private
    GSoc: THGrid;
    HMTrad: THSystemMenu;
   // function LanceImport(Dossier,FileName : string) : variant;
  end ;

Implementation

const NomExec = 'PGIMajVer.exe';
      Tag = '4010';

procedure LanceImport(SocSource,ListSocATraiter,FileName : string);
var LibDoss : string;
begin
  LibDoss := ReadTokenst(ListSocATraiter);
  While  (LibDoss <> '') or (ListSocATraiter <> '') do
  begin
    Application.ProcessMessages;
    if LibDoss <> '' then
    begin
       //FMenuG.LanceDispatchReconnect(LibDoss,nil,False);
       FMenuG.ReconnecteHalley(LibDoss,nil,False);
       //FMenuG.Visible := False;
       AGLIntegreBob (FileName,FALSE,TRUE);
       //FMenuG.FermeSoc;
    end;
    {case iOk of
         0 : stErr:='';
         1 : stErr:='Intégration déjà effectuée';
        -1 : stErr:='Erreur d''écriture dans la table YMYBOBS ';
        -2 : stErr:='Erreur d''intégration dans la fonction AglImportBob';
        -3 : stErr:='Erreur de lecture du fichier BOB.';
        -4 : stErr:='Erreur inconnue.';
    end;  }
    LibDoss := ReadTokenst(ListSocATraiter);
  end;
  FMenuG.ReconnecteHalley(SocSource,nil,False);
  //FMenuG.Visible := True;
end;


procedure TOF_IMPORTOBJETLOT.OnArgument (S : String ) ;
var ListSoc : TListBox;
    iInd : integer;
begin
  Inherited ;
  ListSoc := TListBox(GetControl('LBSOCIETE'));
  GSoc := THGrid(GetControl('GSOC'));
  ChargeDossier(ListSoc.Items,True);
  for iInd := 0 to ListSoc.Items.Count -1 do
  begin
    GSoc.RowCount := GSoc.RowCount + 1;
    GSoc.Cells[1,iInd+1] := ListSoc.Items[iInd];
  end;
  GSoc.RowCount := GSoc.RowCount -1;
  GSoc.ColWidths[0] := -1;
  HMTrad.ResizeGridColumns(GSoc);
end ;

procedure TOF_IMPORTOBJETLOT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_IMPORTOBJETLOT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_IMPORTOBJETLOT.OnUpdate ;
var iInd : integer;
    FileName,SocSource,ListSocATraiter : string;
begin
  Inherited ;
  FileName := GetControlText('FICIMPORT');
  if FileName = '' then
  begin
    PGIBox('Veuillez sélectionner un fichier');
    SetFocusControl('FICIMPORT');
    exit;
  end;
  if GSoc.nbSelected = 0 then
  begin
    PGIBox('Vous devez sélectionner au moins une société');
    SetFocusControl('GSOC');
    exit;
  end;

  if PGIAsk('Confirmez-vous le traitement de ce fichier ?') <> mrYes then exit;

  ListSocATraiter := '';
  for iInd := 1 to GSoc.RowCount -1 do
    if GSoc.IsSelected(iInd) then ListSocATraiter := ListSocATraiter + ';' + GSoc.Cells[1,iInd];
  SocSource := V_PGI.CurrentAlias;
  //FMenuG.FermeSoc;
  //FMenuG.Visible := False;
  LanceImport(SocSource,ListSocATraiter,FileName);
  //FMenuG.Visible := True;
  LastError:=0;
end ;

procedure TOF_IMPORTOBJETLOT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_IMPORTOBJETLOT.OnClose ;
begin
  Inherited ;

end ;

procedure TOF_IMPORTOBJETLOT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_IMPORTOBJETLOT.OnCancel () ;
begin
  Inherited ;
end ;



Initialization
  registerclasses ( [ TOF_IMPORTOBJETLOT ] ) ;
end.
