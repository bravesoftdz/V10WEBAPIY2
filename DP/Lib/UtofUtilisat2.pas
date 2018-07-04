{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 18/09/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : YYUTILISAT2_MUL ()
Mots clefs ... : TOF;UTILISAT2
*****************************************************************}
Unit UtofUtilisat2 ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     eMul,
     MaineAGL,
{$ELSE}
     db,
     dbtables,
     mul,
     FE_Main,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     HDB, HTB97, AGLInit, FileCtrl ;

Type
  TOF_UTILISAT2 = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  private
    sOnglet : String;
    procedure FListe_OnDblClick(Sender : TObject);
    procedure BInsert_OnClick(Sender : TObject);

    // $$$ JP 12/04/06 - assitant droits agenda
    procedure OnClickAssistDroitsAgenda (Sender:TObject);
  end ;


/////////// IMPLEMENTATION ////////////
Implementation

uses EntDP
{$ifdef BUREAU}
, galAssistDroitsAgenda
{$endif bureau}
;

procedure TOF_UTILISAT2.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_UTILISAT2.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_UTILISAT2.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_UTILISAT2.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_UTILISAT2.OnArgument (S : String ) ;
// reçoit en argument l'onglet à afficher par défaut
begin
  Inherited ;
  sOnglet := '';
  if (S='PINTERNET') or (S='PDROITAGENDA') then sOnglet := S;

  // $$$ JP 12/04/06 - assistant droits agenda
  if S='PDROITAGENDA' then
  begin
       with GetControl ('BASSISTDROITSAGENDA') as TToolBarButton97 do
       begin
            Visible := TRUE;
            OnClick := OnClickAssistDroitsAgenda;
       end;
  end;

 //--- CAT 21/01/2008 FQ 11815 : Suppression du bouton nouveau si on n'est pas administrateur
 if not (V_PGI.SuperViseur) then
  GetControl('BINSERT').Visible:=False;

{$IFDEF EAGLCLIENT}
  THGrid(GetControl('FLISTE')).OnDblClick := FListe_OnDblClick;
{$ELSE}
  THDBGrid(GetControl('FLISTE')).OnDblClick := FListe_OnDblClick;
{$ENDIF}
  TToolBarButton97(GetControl('BINSERT')).OnClick := BInsert_OnClick;
end ;

procedure TOF_UTILISAT2.OnClose ;
// Création des répertoires par utilisateur, pour PGIScan
{$IFNDEF EAGLCLIENT}
var tobUsers : TOB;
    racine : String;
    i : Integer;
{$ENDIF}
begin
  Inherited ;
{$IFDEF EAGLCLIENT}
  // #### A FAIRE eAGL si ça a un sens !
{$ELSE}
  if VH_DP.LeServeur='' then exit;
  racine := '\\'+VH_DP.LeServeur+'\pgiscan$';

  if not DirectoryExists(racine) then exit;

  tobUsers := Tob.Create('__UTILISAT_', nil, -1);
  tobUsers.LoadDetailFromSQL('SELECT * FROM UTILISAT');
  tobUsers.SaveToBinFile(racine+'\PGIScanner.tob', False, True, True, True);
  for i:=0 to tobUsers.Detail.Count-1 do
    ForceDirectories(racine + '\' + tobUsers.Detail[i].GetValue('US_UTILISATEUR'));
  tobUsers.Free;
{$ENDIF}
end ;

procedure TOF_UTILISAT2.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_UTILISAT2.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_UTILISAT2.FListe_OnDblClick(Sender : TObject);
begin
  AglLanceFiche('YY','YYUTILISAT','',GetField('US_UTILISATEUR'),
    'ACTION=MODIFICATION;'+sOnglet);
  AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
end ;

procedure TOF_UTILISAT2.Binsert_OnClick(Sender : TObject);
begin
  AglLanceFiche('YY','YYUTILISAT','','','ACTION=CREATION;'+sOnglet);
  AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
end ;

// $$$ JP 12/04/06 - assistant droits agenda
procedure TOF_UTILISAT2.OnClickAssistDroitsAgenda (Sender:TObject);
begin
{$ifdef BUREAU}
     DoAssistDroitsAgenda;
{$endif BUREAU}
end;


Initialization
  registerclasses ( [ TOF_UTILISAT2 ] ) ;
end.

