unit UtilPhases;

interface
uses {$IFDEF VER150} variants,{$ENDIF} HEnt1,
     UTOB,
     Ent1,
     LookUp,
     SysUtils,
     UtilPGI,
     AGLInit,
     HCtrls,
{$IFNDEF EAGLCLIENT}
     FE_Main,
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     Maineagl,
     eMul,
{$ENDIF}
     UtilConso,
     EntGC, Classes, HMsgBox,
     ParamSoc,UPlannifchUtil;
type

TGestionPhase = class
  private
    TOBPhases,CurrentPhase,TOBPhasesDet,TOBCompteur : TOB;
    IndiceCurrent,NivCurrent : integer;
    TOBPiece : TOB;
    GestionConso : TGestionConso;
    GestionConsommation : TTraitConso;
    procedure initphases;
    function GetPhase(TOBL: TOB; Niveau: integer): integer;
    function EncodePhase  : string;
    procedure AddSupPhases(TobPhase : TOB);
    procedure DefiniCOmpteur(Niveau: integer);
    procedure InitCompteur;
  public

    constructor create;
    destructor destroy ; override;
    property Piece : TOB read TOBpiece write TOBpiece;
    procedure clear;
    procedure AjoutePhase (TOBL : TOB);
    procedure RemontePhase;
    procedure AssocieALaPhase (TOBL : TOB);
    procedure CreeAssociation (TOBL : TOB ; Nummouv : double = 0);
    procedure Insere;
    procedure GenerelesPhases (TOBGenere,TOBCONSODEL : TOB;TransfoPiece,DuplicPiece,regroup : boolean;Action : TactionFiche);
    procedure AssocieAuxPhases (TOBGenere,TOBCONSODEL : TOB;TransfoPiece,DuplicPiece,Regroup : boolean;Action : TactionFiche);
    procedure MemoriseLienReception(TOBL: TOB;TOBReception : TOB);
    procedure RecupereReceptions (TOBL : TOB);
    procedure InitialisationLigne (TOBL : TOB);
    procedure InitReceptions;
    procedure recupereReceptionsHorsLien(TOBpiece,TOBAffaire: TOB);
    function NbrReceptionsHorsLien : integer;
    function GetTOBrecepHLien : TOB;
    procedure DefiniReceptionsFromHlienAssocie(TOBL, TOBDetRec: TOB);
    function GetQteLivrable(TOBL: TOB ;  WithGeneration : boolean= false): double;
    procedure RepertorieReceptionsHLiensFromLIV(TOBL: TOB;Affaire: String);
	end;
//var NumMouv  : Double;
procedure AnnulePhases(TOBPiece : TOB; GestionConsommation : TTraitConso = TTcoDelete);
function ISPieceGeneratricePhase (Naturepiece : string) : boolean;
procedure ConstitueLaTOBTrieSurphase ( LaSource,ladestination: TOB );
function GetLibellePhase (TOBL : TOB) : string; overload;
function GetLibellePhase (Chantier,Phase : string) : string; overload;
function GetLibelleChantier (TOBL : Tob) : string;
function ControlePiecePhases (TOBPiece : TOB; var MessageRet : string) : boolean;
function IsExistPhases (Chantier : string) : boolean;
procedure ConstitueStructPhases(TOBPiece,TOBSTP : TOB);
procedure EnvoieInfoPhaseSurLigne(TOBSTP : TOB);

implementation
uses factutil;

procedure AnnuleLignesPhases(TOBPiece : TOB);
var SQL : String;
begin
  Sql := 'DELETE FROM LIGNEPHASES WHERE BLP_NATUREPIECEG="'+
          TOBPiece.getvalue('GP_NATUREPIECEG')+'" AND BLP_SOUCHE="'+
          TOBPiece.getvalue('GP_SOUCHE')+'" AND BLP_NUMERO='+
          IntToStr(TOBPiece.GetValue('GP_NUMERO'))+ ' AND BLP_INDICEG='+
          IntToStr(TOBPiece.GetValue('GP_INDICEG'));
  ExecuteSQL (SQL);
end;

procedure AnnulePhases(TOBPiece : TOB; GestionConsommation : TTraitConso = TTcoDelete);
var SQL : String;
begin
  if isPieceGeneratricePhase(TOBpiece.GetValue('GP_NATUREPIECEG')) then
  begin
    Sql := 'DELETE FROM PHASESCHANTIER WHERE BPC_AFFAIRE="'+TOBPiece.getValue('GP_AFFAIRE')+'"';
    ExecuteSQL (SQL);
  end;
  AnnuleLignesPhases (TOBpiece);
  AnnuleConso(TobPiece,GestionConsommation);
end;

function IsExistPhases (Chantier : string) : boolean;
var QQ : Tquery;
begin
	result := false;
	QQ := OpenSQL ('SELECT BPC_PHASETRA FROM PHASESCHANTIER WHERE BPC_AFFAIRE="'+Chantier+'"',True);
  if not QQ.eof then Result := true;
  Ferme (QQ);
end;

function ISPieceGeneratricePhase (Naturepiece : string) : boolean;
begin
  if VH_GC.BTCODESPECIF = '001' then result := (Naturepiece = 'BCE')
                                else result := (Naturepiece = 'PBT');
end;

function ControlePiecePhases (TOBPiece : TOB; var MessageRet : string) : boolean;
var Indice : integer;
    TOBL : TOB;
    Naturepiece : string;
begin
  result := True;
  NaturePiece := TOBPiece.GetValue('GP_NATUREPIECEG');
  if not PieceAutoriseConso (NaturePiece) then exit;
  for Indice := 0 to TOBPiece.detail.count -1 do
  begin
    TOBL := TOBPiece.detail[Indice];
    if (TOBL.GetValue('BCO_INDICE') = 0) or (TOBL.GetValue('GL_AFFAIRE') = '' ) then continue;
    if (TOBL.GetVAlue('BCO_QUANTITE') <> TOBL.GetValue('GL_QTERESTE')) or
       (TOBL.GetVAlue('BCO_TRANSFORME') = 'X') or
       (TOBL.GetValue('BCO_QTEVENTE') > 0) then
    begin
       Messageret := 'Suppression impossible : Cette pièce à été partiellement livrée';
       result := false;
       break;
    end;
  end;
end;

procedure ConstitueLaTOBTrieSurphase ( LaSource,ladestination: TOB );

function GetNiveauPhase(TOBL : TOB) : integer;
begin
  result := 0;
  if TOBL.GetValue('BPC_PN9') > 0 then result := 9
  else if TOBL.GetValue('BPC_PN8') > 0 then result := 8
  else if TOBL.GetValue('BPC_PN7') > 0 then result := 7
  else if TOBL.GetValue('BPC_PN6') > 0 then result := 6
  else if TOBL.GetValue('BPC_PN5') > 0 then result := 5
  else if TOBL.GetValue('BPC_PN4') > 0 then result := 4
  else if TOBL.GetValue('BPC_PN3') > 0 then result := 3
  else if TOBL.GetValue('BPC_PN2') > 0 then result := 2
  else if TOBL.GetValue('BPC_PN1') > 0 then result := 1;
end;

function rechercheMere(TOBL,LaDestination : TOB; Niveau : integer) : TOB;
begin
	result := nil;
  if niveau = 1 then
  begin
    result := Ladestination;
  end else if niveau = 2 then
  begin
    result := LaDestination.findFirst(['BPC_PN1','BPC_PN2','BPC_PN3',
                                       'BPC_PN4','BPC_PN5','BPC_PN6',
                                       'BPC_PN7','BPC_PN8','BPC_PN9'],
                                       [TOBL.GetValue('BPC_PN1'),0,0,0,0,0,0,0,0],true);
  end else if niveau = 3 then
  begin
    result := LaDestination.findFirst(['BPC_PN1','BPC_PN2','BPC_PN3',
                                       'BPC_PN4','BPC_PN5','BPC_PN6',
                                       'BPC_PN7','BPC_PN8','BPC_PN9'],
                                       [TOBL.GetValue('BPC_PN1'),TOBL.GetValue('BPC_PN2'),
                                        0,0,0,0,0,0,0],true);
  end else if niveau = 4 then
  begin
    result := LaDestination.findFirst(['BPC_PN1','BPC_PN2','BPC_PN3',
                                       'BPC_PN4','BPC_PN5','BPC_PN6',
                                       'BPC_PN7','BPC_PN8','BPC_PN9'],
                                       [TOBL.GetValue('BPC_PN1'),TOBL.GetValue('BPC_PN2'),
                                        TOBL.GetValue('BPC_PN3'),0,0,0,0,0,0],true);
  end else if niveau = 5 then
  begin
    result := LaDestination.findFirst(['BPC_PN1','BPC_PN2','BPC_PN3',
                                       'BPC_PN4','BPC_PN5','BPC_PN6',
                                       'BPC_PN7','BPC_PN8','BPC_PN9'],
                                       [TOBL.GetValue('BPC_PN1'),TOBL.GetValue('BPC_PN2'),
                                        TOBL.GetValue('BPC_PN3'),TOBL.GetValue('BPC_PN4'),
                                        0,0,0,0,0],true);
  end else if niveau = 6 then
  begin
    result := LaDestination.findFirst(['BPC_PN1','BPC_PN2','BPC_PN3',
                                       'BPC_PN4','BPC_PN5','BPC_PN6',
                                       'BPC_PN7','BPC_PN8','BPC_PN9'],
                                       [TOBL.GetValue('BPC_PN1'),TOBL.GetValue('BPC_PN2'),
                                        TOBL.GetValue('BPC_PN3'),TOBL.GetValue('BPC_PN4'),
                                        TOBL.GetValue('BPC_PN5'),0,0,0,0],true);
  end else if niveau = 7 then
  begin
    result := LaDestination.findFirst(['BPC_PN1','BPC_PN2','BPC_PN3',
                                       'BPC_PN4','BPC_PN5','BPC_PN6',
                                       'BPC_PN7','BPC_PN8','BPC_PN9'],
                                       [TOBL.GetValue('BPC_PN1'),TOBL.GetValue('BPC_PN2'),
                                        TOBL.GetValue('BPC_PN3'),TOBL.GetValue('BPC_PN4'),
                                        TOBL.GetValue('BPC_PN5'),TOBL.GetValue('BPC_PN6'),
                                        0,0,0],true);
  end else if niveau = 8 then
  begin
    result := LaDestination.findFirst(['BPC_PN1','BPC_PN2','BPC_PN3',
                                       'BPC_PN4','BPC_PN5','BPC_PN6',
                                       'BPC_PN7','BPC_PN8','BPC_PN9'],
                                       [TOBL.GetValue('BPC_PN1'),TOBL.GetValue('BPC_PN2'),
                                        TOBL.GetValue('BPC_PN3'),TOBL.GetValue('BPC_PN4'),
                                        TOBL.GetValue('BPC_PN5'),TOBL.GetValue('BPC_PN6'),
                                        TOBL.GetValue('BPC_PN7'),0,0],true);
  end else if niveau = 9 then
  begin
    result := LaDestination.findFirst(['BPC_PN1','BPC_PN2','BPC_PN3',
                                       'BPC_PN4','BPC_PN5','BPC_PN6',
                                       'BPC_PN7','BPC_PN8','BPC_PN9'],
                                       [TOBL.GetValue('BPC_PN1'),TOBL.GetValue('BPC_PN2'),
                                        TOBL.GetValue('BPC_PN3'),TOBL.GetValue('BPC_PN4'),
                                        TOBL.GetValue('BPC_PN5'),TOBL.GetValue('BPC_PN6'),
                                        TOBL.GetValue('BPC_PN7'),TOBL.GetValue('BPC_PN8'),0],true);
  end;
end;

var indice,Niveau : integer;
    TOBL,TOBMere,UneTOB : TOB;
begin
  for Indice := 0 to laSource.detail.count -1 do
  begin
    TOBL := LaSource.detail[indice];
    Niveau := GetNiveauPhase(TOBL);
    if Niveau > 0 then
    begin
      TOBmere := rechercheMere (TOBL,LaDestination,Niveau);
      if TOBmere <> nil then
      begin
        UneTob := TOB.Create ('PHASESCHANTIER',TOBmere,-1);
        UneTOB.Dupliquer (TOBL,true,true);
      end;
    end;
  end;
end;

function GetLibelleChantier (TOBL : TOB) : string;
var QQ 				: TQuery;
    Sql 			: String;
    Chantier	: string;
begin
  if (TOBL.FieldExists('LIBELLECHANTIER')) and (TOBL.GetValue('LIBELLECHANTIER') <> '') then
  begin
  	result := TOBL.GetValue('LIBELLECHANTIER');
    exit;
  end;
  result := '';
  chantier := TOBL.GetValue('BCO_AFFAIRE');
  if chantier = '' then exit;

  Sql := 'SELECT AFF_AFFAIRE0, AFF_ETATAFFAIRE, AFF_AFFAIREINIT, AFF_LIBELLE, AFF_TIERS, T_LIBELLE ';
  SQL := SQL + ' FROM AFFAIRE LEFT JOIN TIERS ON AFF_TIERS=T_TIERS WHERE AFF_AFFAIRE="'+Chantier+'"';
  QQ := OpenSql (SQL,true);

  if not QQ.eof then
  begin
    //result := QQ.findfield('AFF_LIBELLE').asString ;
    if TOBL.FieldExists('LIBAFFAIRE') then
      TOBL.PUTVALUE('LIBAFFAIRE', QQ.findfield('AFF_LIBELLE').asString)
    else
      TOBL.AddChampSupValeur('LIBAFFAIRE', QQ.findfield('AFF_LIBELLE').asString);
    //
    TOBL.PutValue('BCO_AFFAIRE0', QQ.findfield('AFF_AFFAIRE0').asString);

    // Modified by f.vautrain 16/10/2017 10:24:56 - FS#2746 - VEODIS - En Saisie conso la modification du contrat n'est pas prise en compte
    If TOBL.GetValue('BCO_AFFAIRESAISIE') = '' then TOBL.PutValue('BCO_AFFAIRESAISIE', QQ.findfield('AFF_AFFAIREINIT').asString);
    //
    //
    TOBL.PUTVALUE('ETATAFFAIRE', QQ.findfield('AFF_ETATAFFAIRE').asString);
    TOBL.PUTVALUE('TIERS', QQ.findfield('AFF_TIERS').asString);
    TOBL.PUTVALUE('LIBTIERS', QQ.findfield('T_LIBELLE').asString);
    //
  end;
  
  ferme (QQ);

end;

function GetLibellePhase (TOBL : TOB) : string;
var QQ : TQuery;
    Sql : String;
begin
  if (TOBL.FieldExists('LIBELLEPHASE')) and (TOBL.GetValue('LIBELLEPHASE') <> '') then
  begin
  	result := TOBL.GetValue('LIBELLEPHASE');
    exit;
  end;
  result := GetLibellePhase(TOBL.GetValue('BCO_AFFAIRE'),TOBL.GetValue('BCO_PHASETRA'));
end;

function GetLibellePhase (Chantier,Phase : string) : string;
var QQ : TQuery;
    Sql : String;
begin
  result := traduireMemoire('Non définie');
  if Phase = '' then exit;
  Sql := 'SELECT BPC_LIBELLE FROM PHASESCHANTIER WHERE BPC_AFFAIRE="'+
          Chantier+'" AND BPC_PHASETRA="'+phase+'"';
  QQ := OpenSql (SQL,true);
  if not QQ.eof then
  begin
    result := QQ.findfield('BPC_LIBELLE').asString ;
  end;
  ferme (QQ);
end;

{ gestionPhase }

procedure TgestionPhase.AjoutePhase(TOBL: TOB);
var TOBphase,TOBLast : TOB;
begin
  TOBLast := nil;
  if IndiceCurrent >= 0 then TOBLast := CurrentPhase;
  inc(NivCurrent);

  TOBphase := TOB.create ('PHASESCHANTIER',TOBPhases,-1);
  definiCompteur (nivcurrent);
  if TOBlast <> nil then
  begin
    TOBPhase.putValue('BPC_PN1',TobLast.getValue('BPC_PN1'));
    if NivCurrent >= 2 then TOBPhase.putValue('BPC_PN2',TobLast.getValue('BPC_PN2'));
    if NivCurrent >= 3 then TOBPhase.putValue('BPC_PN3',TobLast.getValue('BPC_PN3'));
    if NivCurrent >= 4 then TOBPhase.putValue('BPC_PN4',TobLast.getValue('BPC_PN4'));
    if NivCurrent >= 5 then TOBPhase.putValue('BPC_PN5',TobLast.getValue('BPC_PN5'));
    if NivCurrent >= 6 then TOBPhase.putValue('BPC_PN6',TobLast.getValue('BPC_PN6'));
    if NivCurrent >= 7 then TOBPhase.putValue('BPC_PN7',TobLast.getValue('BPC_PN7'));
    if NivCurrent >= 8 then TOBPhase.putValue('BPC_PN8',TobLast.getValue('BPC_PN8'));
    if NivCurrent = 9 then TOBPhase.putValue('BPC_PN9',TobLast.getValue('BPC_PN9'));
    if NivCurrent = 9 then TOBPhase.putValue('BPC_PN9',TobL.getValue('GL_NUMORDRE')) else
    if NivCurrent = 8 then TOBPhase.putValue('BPC_PN8',TobL.getValue('GL_NUMORDRE')) else
    if NivCurrent = 7 then TOBPhase.putValue('BPC_PN7',TobL.getValue('GL_NUMORDRE')) else
    if NivCurrent = 6 then TOBPhase.putValue('BPC_PN6',TobL.getValue('GL_NUMORDRE')) else
    if NivCurrent = 5 then TOBPhase.putValue('BPC_PN5',TobL.getValue('GL_NUMORDRE')) else
    if NivCurrent = 4 then TOBPhase.putValue('BPC_PN4',TobL.getValue('GL_NUMORDRE')) else
    if NivCurrent = 3 then TOBPhase.putValue('BPC_PN3',TobL.getValue('GL_NUMORDRE')) else
    if NivCurrent = 2 then TOBPhase.putValue('BPC_PN2',TobL.getValue('GL_NUMORDRE'));
  end else
  begin
    TOBPhase.putValue('BPC_PN1',TobL.getValue('GL_NUMORDRE'));
  end;
  TOBPhase.putValue('BPC_AFFAIRE',TOBL.GetValue('GL_AFFAIRE'));
  TOBPhase.putValue('BPC_LIBELLE',TOBL.GetValue('GL_LIBELLE'));
  TOBPhase.putValue('BPC_PHASETRA',EncodePhase);
  TOBPhase.SetBoolean('BPC_SAISISSABLE',TOBL.GetBoolean('GLC_SAISISSABLE'));
  TOBPhase.SetBoolean('BPC_VISIBLE',TOBL.GetBoolean('GLC_VISIBLE'));
  indiceCurrent := TOBPhases.detail.count -1;
  CurrentPhase := TOBPhase;
end;

procedure TGestionPhase.AssocieALaPhase(TOBL: TOB);
var TOBPhaseDet : TOB;
begin
  if TOBL.GetValue('GL_AFFAIRE')='' then exit; // si pas d'affaire inutile de créer un lien sur une phase
  TOBphaseDet := TOB.create ('LIGNEPHASES',TOBPhasesDet,-1);
  if CurrentPhase <> nil then
  begin
    TOBphaseDet.putValue('BLP_PHASETRA',CurrentPhase.getValue('BPC_PHASETRA'));
  end else TOBphaseDet.putValue('BLP_PHASETRA',''); // association au chantier
  TOBphaseDet.putValue('BLP_NATUREPIECEG',TOBL.getValue('GL_NATUREPIECEG'));
  TOBphaseDet.putValue('BLP_SOUCHE',TOBL.getValue('GL_SOUCHE'));
  TOBphaseDet.putValue('BLP_NUMERO',TOBL.getValue('GL_NUMERO'));
  TOBphaseDet.putValue('BLP_INDICEG',TOBL.getValue('GL_INDICEG'));
  TOBphaseDet.putValue('BLP_NUMORDRE',TOBL.getValue('GL_NUMORDRE'));
end;

constructor TgestionPhase.create;
begin
  GestionConso := TGestionConso.create;
  TOBphases := TOB.create('LES PHASES',nil,-1);
  TOBphasesDet := TOB.create ('THE LIGNES PHASES',nil,-1);
  TOBCompteur := TOB.create ('LE COMPTEUR DE PHASE',nil,-1);
  AddSupPhases (TOBCompteur);
  initphases;
end;

procedure TGestionPhase.DefiniCOmpteur (Niveau : integer);
begin
  if Niveau = 1 then
  begin
    TOBCompteur.putValue ('PN1',TOBCompteur.GetValue('PN1')+1);
    TOBCompteur.putValue ('PN2',0);
    TOBCompteur.putValue ('PN3',0);
    TOBCompteur.putValue ('PN4',0);
    TOBCompteur.putValue ('PN5',0);
    TOBCompteur.putValue ('PN6',0);
    TOBCompteur.putValue ('PN7',0);
    TOBCompteur.putValue ('PN8',0);
    TOBCompteur.putValue ('PN9',0);
  end else
  if Niveau = 2 then
  begin
    TOBCompteur.putValue ('PN2',TOBCompteur.GetValue('PN2')+1);
    TOBCompteur.putValue ('PN3',0);
    TOBCompteur.putValue ('PN4',0);
    TOBCompteur.putValue ('PN5',0);
    TOBCompteur.putValue ('PN6',0);
    TOBCompteur.putValue ('PN7',0);
    TOBCompteur.putValue ('PN8',0);
    TOBCompteur.putValue ('PN9',0);
  end else
  if Niveau = 3 then
  begin
    TOBCompteur.putValue ('PN3',TOBCompteur.GetValue('PN3')+1);
    TOBCompteur.putValue ('PN4',0);
    TOBCompteur.putValue ('PN5',0);
    TOBCompteur.putValue ('PN6',0);
    TOBCompteur.putValue ('PN7',0);
    TOBCompteur.putValue ('PN8',0);
    TOBCompteur.putValue ('PN9',0);
  end else
  if Niveau = 4 then
  begin
    TOBCompteur.putValue ('PN4',TOBCompteur.GetValue('PN4')+1);
    TOBCompteur.putValue ('PN5',0);
    TOBCompteur.putValue ('PN6',0);
    TOBCompteur.putValue ('PN7',0);
    TOBCompteur.putValue ('PN8',0);
    TOBCompteur.putValue ('PN9',0);
  end else
  if Niveau = 5 then
  begin
    TOBCompteur.putValue ('PN5',TOBCompteur.GetValue('PN5')+1);
    TOBCompteur.putValue ('PN6',0);
    TOBCompteur.putValue ('PN7',0);
    TOBCompteur.putValue ('PN8',0);
    TOBCompteur.putValue ('PN9',0);
  end else
  if Niveau = 6 then
  begin
    TOBCompteur.putValue ('PN6',TOBCompteur.GetValue('PN6')+1);
    TOBCompteur.putValue ('PN7',0);
    TOBCompteur.putValue ('PN8',0);
    TOBCompteur.putValue ('PN9',0);
  end else
  if Niveau = 7 then
  begin
    TOBCompteur.putValue ('PN7',TOBCompteur.GetValue('PN7')+1);
    TOBCompteur.putValue ('PN8',0);
    TOBCompteur.putValue ('PN9',0);
  end else
  if Niveau = 8 then
  begin
    TOBCompteur.putValue ('PN8',TOBCompteur.GetValue('PN8')+1);
    TOBCompteur.putValue ('PN9',0);
  end else
  if Niveau = 9 then
  begin
    TOBCompteur.putValue ('PN9',TOBCompteur.GetValue('PN9')+1);
  end;
end;

destructor TgestionPhase.destroy;
begin
  inherited;
  TOBPhases.free;
  TOBphasesDet.free;
  TOBCompteur.free;
  GestionConso.free;
end;

function TGestionPhase.GetPhase (TOBL : TOB; Niveau : integer) : integer;
begin
  if Niveau = 1 then result := TOBL.GetValue('PN1') else
  if Niveau = 2 then result := TOBL.GetValue('PN2') else
  if Niveau = 3 then result := TOBL.GetValue('PN3') else
  if Niveau = 4 then result := TOBL.GetValue('PN4') else
  if Niveau = 5 then result := TOBL.GetValue('PN5') else
  if Niveau = 6 then result := TOBL.GetValue('PN6') else
  if Niveau = 7 then result := TOBL.GetValue('PN7') else
  if Niveau = 8 then result := TOBL.GetValue('PN8') else
  if Niveau = 9 then result := TOBL.GetValue('PN9') else
  result := 0;
end;


function TGestionPhase.EncodePhase: string;
var P1,P2,P3,P4,P5,P6,P7,P8,P9 : integer;
begin
  result := '';
  P1 := GetPhase(TOBCompteur,1); P2 := GetPhase(TOBCompteur,2); P3 := GetPhase(TOBCompteur,3);
  P4 := GetPhase(TOBCompteur,4); P5 := GetPhase(TOBCompteur,5); P6 := GetPhase(TOBCompteur,6);
  P7 := GetPhase(TOBCompteur,7); P8 := GetPhase(TOBCompteur,8); P9 := GetPhase(TOBCompteur,9);
  if P1 > 0 then result := IntTostr(P1);
  if P2 > 0 then result := result +'.'+IntToStr(P2);
  if P3 > 0 then result := result +'.'+inttoStr(p3);
  if P4 > 0 then result := result +'.'+IntTostr(P4);
  if P5 > 0 then result := result +'.'+IntToStr(P5);
  if P6 > 0 then result := result +'.'+inttoStr(p6);
  if P7 > 0 then result := result +'.'+IntTostr(P7);
  if P8 > 0 then result := result +'.'+IntToStr(P8);
  if P9 > 0 then result := result +'.'+inttoStr(p9);
end;

procedure TgestionPhase.Initphases;
begin
  IndiceCurrent := -1;
  CurrentPhase := nil;
  NivCurrent := 0;
  InitCompteur;
end;

procedure TGestionPhase.InitCompteur;
begin
  TOBcompteur.PutValue('PN1',0);
  TOBcompteur.PutValue('PN2',0);
  TOBcompteur.PutValue('PN3',0);
  TOBcompteur.PutValue('PN4',0);
  TOBcompteur.PutValue('PN5',0);
  TOBcompteur.PutValue('PN6',0);
  TOBcompteur.PutValue('PN7',0);
  TOBcompteur.PutValue('PN8',0);
  TOBcompteur.PutValue('PN9',0);
end;

procedure TGestionPhase.Insere;
begin
  AnnulePhases (TOBPiece,GestionConsommation);
  if (GestionConsommation = TTcoLivraison) then GestionConso.ReajusteReception;
  if not TOBPhasesDet.InsertDB (nil,true) then V_PGI.IOError := OeUnknown;
  if (V_PGI.IOError = OeOk ) and (ISPieceGeneratricePhase(TOBpiece.GetValue('GP_NATUREPIECEG')))  then
  begin
    if not TOBPhases.InsertDB (nil,true) then
    begin
      MessageValid := 'Erreur mise à jour CONSOS (2)';
      V_PGI.IOError := OeUnknown;
    end;
    if (VH_GC.BTCODESPECIF = '001') then
    begin
      if V_PGI.IOError = OeOk then
      begin
        if TOBPiece.GetString('GP_AFFAIRE') <> '' then ExecuteSQL('UPDATE AFFAIRE SET AFF_DATEMODIF="'+USDATETIME(Nowh)+'" WHERE AFF_AFFAIRE="'+TOBPiece.GetString('GP_AFFAIRE')+'"' );
      end;
    end;
  end;
end;

procedure TGestionPhase.RemontePhase;
begin
  dec(IndiceCurrent);
  dec(NivCurrent);
  if (IndiceCurrent >= 0) and (NivCurrent > 0) then
  begin
    CurrentPhase := TOBPhases.detail[IndiceCurrent]
  end else
  begin
    CurrentPhase := nil;
    IndiceCurrent := -1;
    NivCurrent := 0;
  end;
end;

procedure TGestionPhase.AddSupPhases(TobPhase : TOB);
begin
  TOBPhase.AddChampSupValeur('PN1',0);
  TOBPhase.AddChampSupValeur('PN2',0);
  TOBPhase.AddChampSupValeur('PN3',0);
  TOBPhase.AddChampSupValeur('PN4',0);
  TOBPhase.AddChampSupValeur('PN5',0);
  TOBPhase.AddChampSupValeur('PN6',0);
  TOBPhase.AddChampSupValeur('PN7',0);
  TOBPhase.AddChampSupValeur('PN8',0);
  TOBPhase.AddChampSupValeur('PN9',0);
end;

procedure TGestionPhase.CreeAssociation (TOBL : TOB ; Nummouv : double = 0);
var UneTob : TOB;
begin
  if TOBL.GetValue('GL_AFFAIRE')='' then exit;
  if not TOBL.FieldExists('BLP_PHASETRA') then TOBL.AddChampSupValeur('BLP_PHASETRA','');
  UneTob := TOB.Create ('LIGNEPHASES',TOBPhasesDet,-1);
  UneTOB.putValue('BLP_NATUREPIECEG',TOBL.GetValue('GL_NATUREPIECEG'));
  UneTOB.putValue('BLP_SOUCHE',TOBL.GetValue('GL_SOUCHE'));
  UneTOB.putValue('BLP_NUMERO',TOBL.GetValue('GL_NUMERO'));
  UneTOB.putValue('BLP_INDICEG',TOBL.GetValue('GL_INDICEG'));
  UneTOB.putValue('BLP_NUMORDRE',TOBL.GetValue('GL_NUMORDRE'));
  UneTOB.putValue('BLP_PHASETRA',TOBL.GetValue('BLP_PHASETRA'));
  UneTOB.putValue('BLP_NUMMOUV',NumMouv);
end;

procedure TgestionPhase.GenerelesPhases (TOBGenere,TOBCONSODEL : TOB;TransfoPiece,DuplicPiece,regroup : boolean;Action : TactionFiche);
var indice : integer;
    TOBL : TOB;
    TypeLigne,TypeArticle : string;
begin
	if not PieceAutoriseConso(TOBgenere.GetValue ('GP_NATUREPIECEG')) then exit;

  if not ISPieceGeneratricePhase(TOBgenere.GetValue ('GP_NATUREPIECEG')) then
  begin
    AssocieAuxPhases (TOBgenere,TOBCONSODEL,TransfoPiece,DuplicPiece,regroup,Action);
    exit;
  end;
// Dans le cas de la pièce génératrice des phase (donc la prévision)
// aucune consommation n'est enregistrée
  piece := TOBgenere;
  TRY
    for indice := 0 to TOBgenere.detail.count -1 do
    begin
      TOBL := TOBgenere.detail[Indice];
      if TOBL.getValue('GL_AFFAIRE')='' then continue;
      TypeLigne :=TOBL.GetValue('GL_TYPELIGNE');
      if copy(TypeLigne,1,2)='DP' then
      begin
        AjoutePhase (TOBL);
      end else if TypeLigne='ART' then
      begin
        TypeArticle := TOBL.GetString('GL_TYPEARTICLE');
        if (Copy(TypeArticle,1,2)='OU') and (TOBgenere.GetValue ('GP_NATUREPIECEG')='BCE') and (VH_GC.BTCODESPECIF = '001') then
        begin
          if TOBL.GetString('TYPETRANSFERT')<> 'X' then
          begin
            AjoutePhase (TOBL);
            RemontePhase;
          end;
        end else
        begin
          AssocieALaPhase (TOBL);
        end;
      end else if copy(TypeLigne,1,2)='TP' then
      begin
        RemontePhase;
      end;
    end;
    Insere;
  FINALLY
    clear;
  END;
end;

procedure TGestionPhase.AssocieAuxPhases (TOBGenere,TOBCONSODEL : TOB;TransfoPiece,DuplicPiece,Regroup : boolean;Action : TactionFiche);
var indice        : integer;
    TOBL          : TOB;
    NumMouv : double;
    Naturepiece : string;
begin
  NumMouv := 0;
  Naturepiece := TOBgenere.GetValue ('GP_NATUREPIECEG');
  if ISPieceGeneratricePhase(Naturepiece) then exit; // protection
//  GestionPhase := TGestionPhase.create;
  piece := TOBgenere;

  if Regroup then GestionConsommation := TtcoRegroup else
  if {(TransfoPiece) and }((Naturepiece = 'BLC') or (Naturepiece = 'LBT')) then GestionConsommation := TtcoLivraison else
  if TransfoPiece then GestionConsommation := TtcoTransform else
  if (DuplicPiece) and (Action = TaModif) then GestionConsommation := TTcoCreat else
  if Action = TaCreat then GestionConsommation := TTcoCreat else
  if Action = TaModif then GestionConsommation := TTcoModif;
  // Cas de la modif de livraison
  if ((Naturepiece = 'BLC') or (Naturepiece = 'LBT')) and (action=TaModif) and (Not Transfopiece) then
  begin
    gestionConso.gestionModifLivraison (TOBGenere,TOBCONSODEL);
    clear;
    exit;
  end;
  // ---

  TRY
    for Indice := 0 to TOBgenere.detail.count -1 do
    begin
      TOBL := TOBgenere.detail[Indice];
      if TOBL.getValue('GL_AFFAIRE') = '' then continue;
      if TOBL.GetValue('GL_QTEFACT') = 0 then continue;
      if (TOBL.GetValue('GL_TYPELIGNE')='ART') then
      begin
        if PieceAutoriseConso(TOBL.GetValue('GL_NATUREPIECEG')) then
        begin
          NumMouv := GestionConso.CreerConso(TOBL, TOBL.GetValue('BLP_PHASETRA'),GestionConsommation);
        end;
        CreeAssociation (TOBL,Nummouv);
      end;
    end;
    Insere;
  FINALLY
    GestionConso.MajTableConso;
    clear;
  END;
end;

procedure TGestionPhase.clear;
begin
  TOBPhases.ClearDetail;
  TOBphasesDet.Cleardetail;
  TOBCompteur.ClearDetail;
  GestionConso.clear;
  initphases;
end;

procedure TGestionPhase.MemoriseLienReception(TOBL: TOB;TOBReception : TOB);
begin
  // passe a ton voi zin zin
 GestionConso.MemoriseLienReception(TOBL,TOBreception);
end;

procedure TGestionPhase.RecupereReceptions(TOBL: TOB);
begin
  // passe a ton voi zin zin
 GestionConso.RecupereReceptions(TOBL);
end;


procedure TGestionPhase.InitialisationLigne(TOBL: TOB);
begin
	if TOBL.FieldExists('BLP_NUMMOUV') then
  begin
    if VarIsNull(TOBL.getValue('BLP_NUMMOUV')) or (VarAsType(TOBL.getValue('BLP_NUMMOUV'), varString) = #0) then
      TOBL.PutValue('BLP_NUMMOUV',0);
    if VarIsNull(TOBL.getValue('BLP_PHASETRA')) or (VarAsType(TOBL.getValue('BLP_PHASETRA'), varString) = #0) then
      TOBL.PutValue('BLP_PHASETRA','');
    if VarIsNull(TOBL.getValue('BCO_LIENTRANSFORME')) or (VarAsType(TOBL.getValue('BCO_LIENTRANSFORME'), varString) = #0) then
      TOBL.PutValue('BCO_LIENTRANSFORME',0);
    if VarIsNull(TOBL.getValue('BCO_INDICE')) or (VarAsType(TOBL.getValue('BCO_INDICE'), varString) = #0) then
      TOBL.PutValue('BCO_INDICE',0);
    if VarIsNull(TOBL.getValue('BCO_TRANSFORME')) or (VarAsType(TOBL.getValue('BCO_TRANSFORME'), varString) = #0) then
      TOBL.PutValue('BCO_TRANSFORME','-');
    if VarIsNull(TOBL.getValue('BCO_QTEVENTE')) or (VarAsType(TOBL.getValue('BCO_QTEVENTE'), varString) = #0) then
      TOBL.PutValue('BCO_QTEVENTE',0);
    if VarIsNull(TOBL.getValue('BCO_QUANTITE')) or (VarAsType(TOBL.getValue('BCO_QUANTITE'), varString) = #0) then
      TOBL.PutValue('BCO_QUANTITE',0);
    if VarIsNull(TOBL.getValue('BCO_LIENVENTE')) or (VarAsType(TOBL.getValue('BCO_LIENVENTE'), varString) = #0) then
      TOBL.PutValue('BCO_LIENVENTE',0);
    if  VarIsNull(TOBL.getValue('BCO_TRAITEVENTE')) or (VarAsType(TOBL.getValue('BCO_TRAITEVENTE'), varString) = #0) then
      TOBL.PutValue('BCO_TRAITEVENTE','-');
    if VarIsNull(TOBL.getValue('BCO_QTEINIT')) or (VarAsType(TOBL.getValue('BCO_QTEINIT'), varString) = #0) then
      TOBL.PutValue('BCO_QTEINIT',0);
    if VarIsNull(TOBL.getValue('BCO_LIENRETOUR')) or (VarAsType(TOBL.getValue('BCO_LIENRETOUR'), varString) = #0) then
      TOBL.PutValue('BCO_LIENRETOUR',0);
  end;
end;

procedure TGestionPhase.InitReceptions;
begin
  // passe a ton voi zin zin
 GestionConso.InitReceptions;
end;

procedure TGestionPhase.recupereReceptionsHorsLien (TOBPiece,TOBAffaire : TOB);
begin
  // passe a ton voi zin zin
 GestionConso.recupereReceptionsHorsLien (TOBPiece,TOBaffaire);
end;

function TGestionPhase.NbrReceptionsHorsLien: integer;
begin
 result := GestionConso.NbrReceptionsHorsLien;
end;

function TGestionPhase.GetTOBrecepHLien: TOB;
begin
 result := GestionConso.GetTOBrecepHLien;
end;

procedure TGestionPhase.DefiniReceptionsFromHlienAssocie (TOBL,TOBDetRec : TOB);
begin
	gestionConso.DefiniReceptionsFromHlienAssocie (TOBL,TOBDetRec);
end;

function TGestionPhase.GetQteLivrable (TOBL : TOB ; WithGeneration : boolean= false) : double;
begin
	result := gestionConso.GetQteLivrable(TOBL,WithGeneration);
end;

procedure TGestionPhase.RepertorieReceptionsHLiensFromLIV (TOBL : TOB ; Affaire : String);
begin
	gestionConso.RepertorieReceptionsHLiensFromLIV (TOBL,Affaire);
end;

procedure ConstitueStructPhases(TOBPiece,TOBSTP : TOB);

  procedure AjouteUnePhase(TOBL,TOBSTP : TOB);
  var TOBP : TOB;
  begin
    TOBP := TOB.Create ('UNE PHASE',TOBSTP,-1);
    TOBP.AddChampSupValeur('REFERENCE',TOBL.GetString('GLC_NUMEROTATION'));
    TOBP.AddChampSupValeur('LIBELLE',TOBL.GetString('GL_LIBELLE'));
    TOBP.AddChampSupValeur('VISIBLE',TOBL.GetBoolean('GLC_VISIBLE'));
    TOBP.AddChampSupValeur('SAISISSABLE',TOBL.GetBoolean('GLC_SAISISSABLE'));
    if TOBL.GetString('GL_TYPEARTICLE')='OUV' then
    begin
      TOBP.AddChampSupValeur('NIVEAU',TOBL.GetInteger('GL_NIVEAUIMBRIC')+1);
    end else
    begin
      TOBP.AddChampSupValeur('NIVEAU',TOBL.GetInteger('GL_NIVEAUIMBRIC'));
    end;
    TOBP.data := TOBL;
  end;

var II : Integer;
    TOBL : TOB;
begin
  for II := 0 to TOBPiece.Detail.count -1 do
  begin
    TOBL := TOBPiece.detail[II];
    if IsDebutParagraphe(TOBL) or (TOBL.GetString('GL_TYPEARTICLE')='OUV') then
    begin
      AjouteUnePhase(TOBL,TOBSTP);    
    end;
  end;
end;

procedure EnvoieInfoPhaseSurLigne(TOBSTP : TOB);
var II : Integer;
    TOBP,TOBL : TOB;
begin
  for II := 0 TO TOBSTP.detail.count -1 do
  begin
    TOBP := TOBSTP.detail[II];
    TOBL := TOB(TOBP.data);
    if TOBL <> NIL then
    begin
      TOBL.SetBoolean('GLC_VISIBLE',TOBP.GetBoolean('VISIBLE'));
      TOBL.SetBoolean('GLC_SAISISSABLE',TOBP.GetBoolean('SAISISSABLE'));
    end;
  end;
end;

end.
