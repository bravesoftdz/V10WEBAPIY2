{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 09/07/2008
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTCHOIXAPPLICAVAN ()
Mots clefs ... : TOF;BTCHOIXAPPLICAVAN
*****************************************************************}
Unit BTCHOIXAPPLICAVAN_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     maineagl,
     eMul,
{$ENDIF}
     forms,
     grids,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     AglInit,
     m3FP,
     UTOF,
     graphics,
     windows,
     UTOB ;

Type
  TOF_BTCHOIXAPPLICAVAN = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    TACCEPT : TCheckBox;
    GS : THGrid;
    NbCols : integer;
    TOBCHOIX : TOB;
    avancement : boolean;
    ListeChamps: string;
    LIGNE,LIBELLE,ANOMALIE,OK : integer;
    procedure AfficheLaGrille;
    procedure AfficheLaLigne(Ligne: integer);
    procedure DefiniAttributCol(Nom: string; Colonne: integer; LaLargeur,Lalignement, Letitre, LeNC: string);
    procedure DefinieGrille;
    function GetTOBLigneAV(Arow: integer): TOB;
    procedure GSDBlCLick(Sender: Tobject);
    procedure GetComponents;
    procedure AcceptGlobaleClick (Sender : Tobject);
    procedure SetAcceptGlobale (Etat : boolean);

  end ;

procedure DefiniChoixApplication (TOBDecisionL : TOB);

Implementation

procedure DefiniChoixApplication (TOBDecisionL : TOB);
begin
  TheTOB := TOBDecisionL;
  AGLLanceFiche ('BTP','BTCHOIXAPPLICAVAN','','','');
  TheTOB := nil;
end;

procedure TOF_BTCHOIXAPPLICAVAN.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTCHOIXAPPLICAVAN.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTCHOIXAPPLICAVAN.OnUpdate ;
begin
  Inherited ;
  laTOB.putvalue('OKSAISIE','X');
end ;

procedure TOF_BTCHOIXAPPLICAVAN.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTCHOIXAPPLICAVAN.OnArgument (S : String ) ;
begin
  Inherited ;
  TOBCHOIX :=LaTOB;
  avancement := (TOBCHoix.getValue('AVANCEMENT')='X');
  GetComponents;
  TACCEPT.onclick := AcceptGlobaleClick;
  GS.OnDblClick  := GSDblClick;
  DefinieGrille;
  AfficheLaGrille;
end ;

procedure TOF_BTCHOIXAPPLICAVAN.GetComponents ;
begin
  GS := THgrid(GetControl ('GS'));
  TACCEPT := TCheckBox(GetControl('TACCEPT'));
end;

procedure TOF_BTCHOIXAPPLICAVAN.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTCHOIXAPPLICAVAN.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTCHOIXAPPLICAVAN.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTCHOIXAPPLICAVAN.DefiniAttributCol(Nom: string; Colonne: integer; LaLargeur,Lalignement, Letitre, LeNC: string);
var FF: string;
    Obli,OkLib,OkVisu,OkNulle,OkCumul,Sep : boolean;
    dec : integer;
begin

  TransAlign(Lalignement,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
  GS.cells[Colonne,0] := leTitre;
  // Alignement dans la colonne
  if copy(LAlignement,1,1)='G' then GS.ColAligns[Colonne] := taLeftJustify
  else if copy(LAlignement,1,1)='D' then GS.ColAligns[Colonne] := taRightJustify
  else if copy(LAlignement,1,1)='C' then GS.ColAligns[Colonne] := taCenter;
  GS.ColWidths[Colonne] := strtoint(lalargeur)*GS.Canvas.TextWidth('W');
  if OkLib then
  begin
    GS.ColFormats[Colonne] := 'CB=' + Get_Join(Nom);
  end else if (Dec<>0) or (Sep) then
  begin
    GS.ColFormats[Colonne] := FF ;
  end else if Nom='OK' then
  begin
    GS.ColTypes[colonne] := 'B';
    GS.ColFormats[colonne] := IntToStr( Byte( csCheckBox ) );
  end;

  if LeNc = '0' then GS.ColEditables [Colonne]:= false;
end;

procedure TOF_BTCHOIXAPPLICAVAN.DefinieGrille;
var lesElements,Lelement,LaLargeur,Lalignement,LeTitre,LeNC : string;
    colonne : integer;
    flargeur,Falignement,FTitre,NC : string;
begin
  ListeChamps := 'LIGNE;LIBELLE;ANOMALIE;OK';
  flargeur := '2;19;19;7;';
  Falignement := 'G.0O ---;G.0  ---;G.0  ---;C.0  ---;';
  FTitre := 'N°;Désignation;Anomalie;Choix;';
  NC := '0;0;0;1;'; //definition des zones saisissable ou non

  NbCols := 4 ;// Nombre de colonnes dans la saisie
  GS.ColCount := NbCols;
  lesElements := ListeChamps;
  //
  Lelement := ReadtokenSt(lesElements);
  LaLargeur := ReadtokenSt(fLargeur);
  Lalignement := ReadtokenSt(Falignement);
  LeTItre := readTokenSt(Ftitre);
  LeNC := readTokenSt(NC);
  Colonne := 0;
  Repeat
    if Lelement = '' then break;
    if Lelement = 'LIGNE' then
    begin
      LIGNE := Colonne;
      DefiniAttributCol (Lelement,Colonne,LaLargeur,Lalignement,Letitre,LeNC);
    end else if Lelement = 'LIBELLE' then
    begin
      LIBELLE := Colonne;
      DefiniAttributCol (Lelement,Colonne,LaLargeur,Lalignement,Letitre,LeNC);
    end else if Lelement = 'ANOMALIE' then
    begin
      ANOMALIE := Colonne;
      DefiniAttributCol (Lelement,Colonne,LaLargeur,Lalignement,Letitre,LeNC);
    end else if Lelement = 'OK' then
    begin
      OK := Colonne;
      DefiniAttributCol (Lelement,Colonne,LaLargeur,Lalignement,Letitre,LeNC);
    end;
    //
    Lelement := ReadtokenSt(lesElements);
    LaLargeur := ReadtokenSt(fLargeur);
    Lalignement := ReadtokenSt(Falignement);
    LeTItre := readTokenSt(Ftitre);
    LeNC := readTokenSt(NC);
    inc(colonne);
    //
  until lelement = '';
end;

procedure TOF_BTCHOIXAPPLICAVAN.AfficheLaGrille;
var Indice : integer;
begin
  GS.RowCount := TOBCHOIX.detail.count+1;
  For Indice := 0 to TOBCHOIX.detail.count -1 do
  begin
    AfficheLaLigne (Indice+1);
  end;
end;

procedure TOF_BTCHOIXAPPLICAVAN.AfficheLaLigne (Ligne : integer);
var TOBL : TOB;
begin
  TOBL := GetTOBLigneAV (Ligne); if TOBl = nil then exit;
  TOBL.PutLigneGrid (GS,Ligne,false,false,ListeChamps);
end;

function TOF_BTCHOIXAPPLICAVAN.GetTOBLigneAV(Arow: integer): TOB;
begin
  result := nil;
  if ARow > TOBCHOIX.detail.count then exit;
  result := TOBCHOIX.detail[Arow-1];
end;


procedure TOF_BTCHOIXAPPLICAVAN.GSDBlCLick(Sender: Tobject);
var TOBL : TOB;
begin
  TOBL := GetTOBLigneAV (GS.row); if TOBL = nil then exit;
    //
  if TOBL.GetValue('OK')='-' then TOBL.putValue('OK','X') else TOBL.putValue('OK','-');
  AfficheLaLigne (GS.row);
end;

procedure TOF_BTCHOIXAPPLICAVAN.AcceptGlobaleClick(Sender: Tobject);
begin
  SetAcceptGlobale (TACCEPT.checked);
  AfficheLaGrille;  
end;

procedure TOF_BTCHOIXAPPLICAVAN.SetAcceptGlobale(Etat: boolean);
var Indice : integer;
    TOBL : TOB;
begin
  For Indice := 0 to TOBCHOIX.detail.count -1 do
  begin
    TOBL := TOBCHOIX.detail[Indice]; if TOBL = nil then exit;
    if Etat then TOBL.putValue('OK','X') else TOBL.putValue('OK','-');
  end;
end;

Initialization
  registerclasses ( [ TOF_BTCHOIXAPPLICAVAN ] ) ;
end.
