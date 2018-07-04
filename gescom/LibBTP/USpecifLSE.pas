unit USpecifLSE;

interface

uses sysutils,classes,windows,messages,hmsgbox,
     HEnt1,Ent1,EntGC,UtilPGI,UTOB,
     FactTOB, FacTPiece, FactArticle,
		 UtilXlsBTP,ed_tools,variants,
     HDB,Hqry,Hctrls, CBPPath,
     SaisUtil, Forms, Mul,Controls,
     M3FP,AglInit,
{$IFDEF EAGLCLIENT}
  maineagl,
{$ELSE}
  DBCtrls, Db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} fe_main,UserChg,
{$ENDIF}

{$IFDEF BTP}
 BTPUtil, CalcOLEGenericBTP, HTB97,HRichOle,
{$ENDIF}
 ParamSoc;


type

TTypePiece = (ttypDocMat,TTypDocPrest,TTypDocCont,TTypOsef);
TModegener = (tMgPropal,TmgConvention);

Tgenerepropal = class
	private
  	fnaturepiece : string;
    fSouche : string;
    fNumero : integer;
    fIndiceg : integer;
    fModeGener : TmodeGener;
    //
    fdocExcel : string;
    fWinExcel: OleVariant;
    fWorkBook : Variant;
    //
		fInternalWindow : TToolWindow97;
    fTexte : THRichEditOLE;
    fblocNote : THRichEditOLE;
    ISBlocVide : boolean;
  	//
    fTOBContact : TOB;
  	procedure AddChampsDetail(OneTOB : TOB);
    procedure DefiniEntetePropale (TOBpiece,TOBarchitect : TOB);
    function AjouteDetail(TOBarchitect: TOb; NomEnsemble: string; TOBL: TOB): TOB;
    procedure AjouteFraisPrepMat(TOBarchitect: TOb; TOBL: TOB);
    procedure AjouteRemise(TOBarchitect: TOb; NomEnsemble: string; TOBL: TOB);
    function GetDocument (TOBPiece,TOBComment : TOB) : boolean;
    procedure ConstitueStructure (TOBPiece,TOBarchitect : TOB);
    procedure ConstitueFichierExcel(TOBarchitect,TOBComment : TOB);
    procedure EnregistreFichierExcel(TOBArchitect : TOB);
    procedure AddLesChampsSupArchi(TOBarchitect : TOB);
    function AjouteBlocNote(TOBarchitect: TOB; NomEnsemble: string; TOBL: TOB): TOB;
    function IsBlocNoteVide(texte: THRichEditOLE;BlocNote: String): boolean;
    procedure SetModeGener(const Value: TModegener);

  public
  	constructor create;
    destructor destroy; override;
    property Typegeneration : TModegener read fModeGener write SetModeGener;
    procedure TraiteDocument;


end;


TnewPiece = class (TObject)
	private
    fDEV : Rdevise;
  	fTOBPiece : TOB;
    fTOBAdresses : TOB;
    fTOBOuvrages : TOB;
    fTOBOuvragesP : TOB;
    fTOBBases : TOB;
    fTOBBasesL : TOB;
    fTOBTiers : TOB;
    fTOBArticles : TOB;
    fTOBAffaire : TOB;
    fTOBAffaireP : TOB;
    fTOBFACTAFF : TOB;
    fTOBPieceTrait : TOB;
    fTOBSSTrait : TOB;
    fTOBPorcs,fTOBPieceRG,fTOBBasesRG,fTOBEches : TOB;
    fTOBAffaireAdrInt : TOB;
    fTOBVTECOLL : TOB;
    fType : TTypePiece;
    fnewNum : integer;
    fDateAcc,fDateDebut : TDateTime;
    fOldnaturePieceg,fOldSouche : string;
    fOldNumero, FoldIndiceg : integer;
    //
    function Controleparagraphe(TOBP: TOB; Debut, Niveau: integer;
      var FinParagraphe: integer): boolean;
    procedure DeleteParagraphe(TOBP: TOB; Debut, Niveau: integer);
    procedure NettoieParagraphe(TOBP: TOB);
    procedure AffecteNaturePiece;
    procedure SetDefinitiveNumberAndOthers;
    procedure InitToutModif;
    procedure ConstitueCodeContrat (var AffaireContrat : string);
    procedure AffecteEtatPiece(CodeEtat: string);

  public
  	constructor create;
    destructor destroy; override;
    property TypePiece : TTypePiece read fType write fType;
    property TOBPiece : TOB read fTOBpiece write fTOBPiece;
    property TOBAdresses : TOB read fTOBAdresses write fTOBAdresses;
    property TOBTiers : TOB read fTOBTiers write fTOBTiers;
    property DEV : Rdevise read fDEV write fDEV;
    procedure nettoie;
    procedure ReCalcule;
    procedure SetInfosPiece;
    procedure EcritDocuments;
    procedure SetAffaire ( Affaire: string);
    procedure SetDatePiece (DateCreat : TDateTime);
    procedure ConstitueContrat (AffaireContrat,Affaire : string);
    procedure GenereEcheancesContrat;
    procedure DefiniAffairePiece;
end;

TlisTnewPiece = class (Tlist)
  private
    function GetItems(Indice: integer): TnewPiece;
    procedure SetItems(Indice: integer; const Value: TnewPiece);
  public
    function Add(AObject: TnewPiece): Integer;
    property Items [Indice : integer] : TnewPiece read GetItems write SetItems;
    function findPiece (TypePiece : TTypePiece ): TnewPiece;
    destructor destroy; override;
    procedure clear; override;
end;

TnewPieceGestion = class (Tobject)
	private
  	fnaturepiece : string;
    fSouche : string;
    fNumero : integer;
    fIndiceg : integer;
    fListPiece : TlisTnewPiece; // liste des pieces à générer
    fTOBPiece : TOB;
    fTOBAdresses : TOB;
    fTOBTiers : TOB;
    fDEV : Rdevise;
    fOkTraitement : boolean;
    fAffaire,FaffaireContrat : string;
    fDateAcc,fDateDeb : TdateTime;
    //
    procedure EclateLeDocument;
    procedure EcritlesDocuments;
    procedure BeforeEcriture;
    procedure EcritureDocuments;
    procedure PositionneEtudeEnTermine;
  public
  	constructor create;
    destructor destroy; override;
    property Naturepieceg : string read fnaturepiece write fnaturepiece;
    property Souche : string read fSouche write fSouche;
    property Numero : integer read fNumero  write fNumero;
    property Indiceg : integer read fIndiceg  write fIndiceg;
    property OkTraitement: Boolean read fOkTraitement write fOkTraitement;
    procedure ChargeLeDocument;
    procedure ConstitueDocuments;
    procedure DemandeInfosSupl;


end;

TlistDesPieces = class (Tlist)
  private
    function GetItems(Indice: integer): TnewPieceGestion;
    procedure SetItems(Indice: integer; const Value: TnewPieceGestion);
  public
    function Add(AObject: TnewPieceGestion): Integer;
    property Items [Indice : integer] : TnewPieceGestion read GetItems write SetItems;
    destructor destroy; override;
    procedure clear; override;
end;

procedure LSEGenerepropal (naturePiece,Souche : string; Numero,Indice : integer);
procedure LSEGenereConvention (naturePiece,Souche : string; Numero,Indice : integer);
procedure LSEGenerePiecesGestion (TOBdesPieces : TOB);

implementation

uses UEntCommun,UtilTOBpiece,FactAdresse,FactUtil,FactCalc,FactRG,
		 factureBTP,FactComm,AffaireUtil,AffEcheanceUtil,factSpec,
  TntWideStrings;

procedure LSEGenerePiecesGestion (TOBdesPieces : TOB);
var ThePieceGestion : TnewPieceGestion;
		naturePiece,Souche : string;
    Numero,IndiceG : string;
    Indice : integer;
    TheListPieces : TlistDesPieces;
    TOBD : TOB;
    TheChamps : string;
    AllOk : boolean;
begin
	AllOk := true;
	TheListPieces := TlistDesPieces.create;
  TRY
    for Indice := 0 to TOBDesPieces.detail.count -1 do
    begin
      TOBD := TOBDesPieces.detail[Indice];
      TheChamps := TOBD.getValue('CLEDOC');
      NaturePiece := READTOKENST(TheChamps);
      Souche := READTOKENST(TheChamps);
      Numero := READTOKENST(TheChamps);
      Indiceg := READTOKENST(TheChamps);
      //
      ThePieceGestion := TnewPieceGestion.create;
      ThePieceGestion.Naturepieceg := NaturePiece;
      ThePieceGestion.Souche := Souche;
      ThePieceGestion.Numero := strtoint(Numero);
      ThePieceGestion.Indiceg := strtoint(Indiceg);
      //
      ThePiecegestion.DemandeInfosSupl;
      //
      TheListPieces.Add(ThePiecegestion);
    end;
    for Indice := 0 to TheListPieces.Count -1 do
    begin
      ThePieceGestion := TheListPieces.Items [Indice];
      With ThePieceGestion do
      begin
        TRY
          ConstitueDocuments;
        EXCEPT
    			PgiError ('Des erreurs ce sont produites pendant la validation de l''étude '+
                    inttoStr(ThePieceGestion.Numero)+
          					'. Merci de contacter votre administrateur');
          AllOk := false;
        END;
      end;
    end;
  FINALLY
  	TheListPieces.free;
  END;
  if AllOk then
  begin
  	PgiInfo ('Les documents ont été générés avec succès');
  end;
end;

procedure LSEGenerepropal (naturePiece,Souche : string; Numero,Indice : integer);
var ThePropale : Tgenerepropal;
begin
  ThePropale := Tgenerepropal.create;
  TRY
    With Thepropale do
    begin
      fnaturepiece := Naturepiece;
      fSouche := Souche;
      fNumero := Numero;
      fIndiceg  := Indice;
      Typegeneration := tMgPropal;
      TraiteDocument;
    end;
  FINALLY
  	ThePropale.free;
  END;
end;

procedure LSEGenereConvention (naturePiece,Souche : string; Numero,Indice : integer);
var ThePropale : Tgenerepropal;
begin
  ThePropale := Tgenerepropal.create;
  TRY
    With Thepropale do
    begin
      fnaturepiece := Naturepiece;
      fSouche := Souche;
      fNumero := Numero;
      fIndiceg  := Indice;
      Typegeneration := TmgConvention;
      TraiteDocument;
    end;
  FINALLY
  	ThePropale.free;
  END;
end;


{ Tgenerepropal }

procedure Tgenerepropal.AddChampsDetail(OneTOB: TOB);
begin
	OneTOB.AddChampSupValeur('LIBELLE','');
	OneTOB.AddChampSupValeur('QTE',0);
	OneTOB.AddChampSupValeur('PU',0);
	OneTOB.AddChampSupValeur('MT',0);
	OneTOB.AddChampSupValeur('TYPESERVICE','');
end;

function Tgenerepropal.AjouteDetail(TOBarchitect : TOb;NomEnsemble : string; TOBL : TOB): TOB;
var TOBPere : TOB;
begin
	TOBpere := TOBarchitect.findFirst(['NOMENSEMBLE'],[NomEnsemble],true);
  if TOBPere = nil then
  begin
  	TOBPere := TOB.Create ('UN ENSEMBLE',TOBarchitect,-1);
  	TOBPere.AddChampSupValeur('NOMENSEMBLE',NomEnsemble);
  	TOBPere.AddChampSupValeur('FRAISPREPARATION',0);
  	TOBPere.AddChampSupValeur('POURCENTPREP','');
  	TOBPere.AddChampSupValeur('REMISES',0);
  end;
  result := TOB.create('ONE DETAIL',TOBpere,-1);
  AddChampsDetail(result);
  result.putValue('LIBELLE',TOBL.getValue('GL_LIBELLE'));
  result.putValue('QTE',TOBL.getValue('GL_QTEFACT'));
  result.putValue('PU',TOBL.getValue('GL_PUHTDEV'));
  result.putValue('MT',TOBL.getValue('GL_MONTANTHTDEV'));
  if NomEnsemble = 'ASSISTANCE' then
  begin
  	if Pos('/01',TOBL.getValue('GL_CODEARTICLE')) > 0 then result.putValue('TYPESERVICE','Optimal')
    																							    else result.putValue('TYPESERVICE','Prémium');
  end;
end;

function Tgenerepropal.AjouteBlocNote (TOBarchitect: TOB; NomEnsemble : string;TOBL : TOB) : TOB;
var TOBPere : TOB;
		Indice : integer;
begin
  if not IsBlocNoteVide(fblocNote ,TOBL.getValue('GL_BLOCNOTE')) then
  begin
    TOBpere := TOBarchitect.findFirst(['NOMENSEMBLE'],[NomEnsemble],true);
    if TOBPere = nil then
    begin
      TOBPere := TOB.Create ('UN ENSEMBLE',TOBarchitect,-1);
      TOBPere.AddChampSupValeur('NOMENSEMBLE',NomEnsemble);
      TOBPere.AddChampSupValeur('FRAISPREPARATION',0);
      TOBPere.AddChampSupValeur('POURCENTPREP','');
      TOBPere.AddChampSupValeur('REMISES',0);
    end;

  	for Indice := 0 to fblocNote.lines.Count -1 do
    begin
      result := TOB.create('ONE DETAIL',TOBpere,-1);
      AddChampsDetail(result);
      result.putValue('LIBELLE',fblocNote.Lines.Strings[Indice]);
    end;
  end;
end;

procedure Tgenerepropal.AjouteRemise(TOBarchitect: TOb; NomEnsemble : string; TOBL : TOB);
var TOBPere : TOB;
begin
	TOBpere := TOBarchitect.findFirst(['NOMENSEMBLE'],[NomEnsemble],true);
  if TOBPere = nil then
  begin
  	TOBPere := TOB.Create ('UN ENSEMBLE',TOBarchitect,-1);
  	TOBPere.AddChampSupValeur('NOMENSEMBLE',NomEnsemble);
  	TOBPere.AddChampSupValeur('FRAISPREPARATION',0);
  	TOBPere.AddChampSupValeur('POURCENTPREP','');
  	TOBPere.AddChampSupValeur('REMISES',0);
  end;
  TOBPere.putValue('REMISES',TOBPere.getValue('REMISES')+TOBL.getValue('GL_MONTANTHTDEV'));
end;

procedure Tgenerepropal.AjouteFraisPrepMat(TOBarchitect: TOb; TOBL : TOB);
var TOBC : TOB;
begin
	TOBC := TOBarchitect.findFirst(['NOMENSEMBLE'],['MATERIEL'],true);
  if TOBC <> nil then
  begin
  	TOBC.PutValue('FRAISPREPARATION',TOBC.GetValue('FRAISPREPARATION')+TOBL.getValue('GL_MONTANTHTDEV'));
  	TOBC.PutValue('POURCENTPREP',floattostr(TOBL.getValue('GL_QTEFACT'))+'%');
  end;
end;

procedure Tgenerepropal.ConstitueStructure (TOBPiece,TOBarchitect : TOB);
var indice : integer;
    Famille1,Famille2,Domaine : string;
    Mt : double;
    TOBL : TOB;
begin
	DefiniEntetePropale (TOBPiece,TOBarchitect);
  for Indice := 0 to TOBPiece.detail.count -1 do
  begin
  	TOBL := TOBPiece.detail[Indice];
    //
    Domaine := TOBL.Getvalue('GL_DOMAINE');
    Famille1 := TOBL.Getvalue('GL_FAMILLENIV1');
    Famille2 := TOBL.Getvalue('GL_FAMILLENIV2');
    Mt :=TOBL.Getvalue('GL_MONTANTHTDEV');
    //
    if ((Famille1='PRE') AND (copy(Famille2,1,2)='FD')) then // gestion des frais de déplacement
    begin
    	AjouteDetail (TOBarchitect,'FRAISJOUR',TOBL);
    end else if ((Domaine = 'PRO') AND(Pos(Famille1,'MAT;LBD;LOG;MAI')>0)) OR
       ((Famille1='PRE') AND(Famille2='MAI')) then
    begin
			// Matériels
    	if Mt < 0 then
      begin
    		AjouteRemise (TOBarchitect,'MATERIEL',TOBL);
    		AjouteDetail (TOBarchitect,'MATERIEL',TOBL);
      end else
      begin
    		AjouteDetail (TOBarchitect,'MATERIEL',TOBL);
        AjouteBlocNote (TOBarchitect,'MATERIEL',TOBL);
      end;

    end else if Domaine = 'PON' then
    begin
    	if ((Famille1='PRE') AND (Famille2='INS'))   then
      begin
      	// prestations systeme - réseaux - Logiciels externes
	    	AjouteDetail (TOBarchitect,'PRESTMAT',TOBL);
        AjouteBlocNote (TOBarchitect,'PRESTMAT',TOBL);
      end else if ((Famille1='PRE') AND (Famille2='PAR')) then
      begin
      	// prestations progiciels
	    	AjouteDetail (TOBarchitect,'PRESTPROG',TOBL);
        AjouteBlocNote (TOBarchitect,'PRESTPROG',TOBL);
      end else if ((Famille1='PRE') AND (Famille2='FOR')) then
      begin
        // Formations sur progiciels
        AjouteDetail (TOBarchitect,'FORMATION',TOBL);
        AjouteBlocNote (TOBarchitect,'FORMATION',TOBL);
      end;

    end else if (Domaine = 'REC') then
    begin
    	// Assistance (Contrat)
    	AjouteDetail (TOBarchitect,'ASSISTANCE',TOBL);

    end else if (Famille1='PRE') AND (Famille2='INS')  then
  	begin
    	// Frais de préparation matériels
      AjouteFraisPrepMat (TOBarchitect,TOBL);

    end else if (Famille1='PRO') then
    begin
    	// progiciels LSE
    	if (Mt < 0 )then
      begin
    		AjouteRemise (TOBarchitect,'PROGICIELS',TOBL);
    		AjouteDetail (TOBarchitect,'PROGICIELS',TOBL);
      end else
      begin
      	if (Famille2 = 'LBB') then
        begin
        	TOBarchitect.putValue('TYPEPROGICIEL','BUS');   // défini à Business BTP
        end;
    		AjouteDetail (TOBarchitect,'PROGICIELS',TOBL);
        AjouteBlocNote (TOBarchitect,'PROGICIELS',TOBL);
      end;
    end;
  end;
end;

procedure Tgenerepropal.ConstitueFichierExcel(TOBarchitect,TOBComment : TOB);
var WinNew : boolean;
		Indice,Ind,PosPrestations,PosTexte : integer;
    TOBD,TOBLL : TOB;

begin
	PosPrestations := 0;
  PosTexte := 0;
  if not OpenExcel (true,fWinExcel,WinNew) then
  begin
    PGIBox (TraduireMemoire('Liaison Excel impossible'),'Export EXCEL');
    FDocExcel := '';
    exit;
  end;
  fWorkBook := OpenWorkBook (FDocExcel ,fWinExcel);
  // Particularité
  ExcelRangeValue(fWorkBook,'Synthese','FRAISJOUR','Forfait');
  // --
  if not ISBlocVide then
  begin
  	for Indice := 0 to fTexte.lines.Count -1 do
    begin
    	Inc(PosTexte);
      if Postexte < 25 then
      begin
  			ExcelRangeValue(fWorkBook,'Rappel','TEXTEDEBUT',ftexte.lines.Strings [Indice],PosTexte);
      end else break;
    end;
  end;
  // --
  ExcelRangeValue(fWorkBook,'Proposition','SIRETCLI',TOBarchitect.getValue('SIRETCLI'));
  ExcelRangeValue(fWorkBook,'Proposition','NUMEROCLI',TOBarchitect.getValue('NUMEROCLI'));
  ExcelRangeValue(fWorkBook,'Proposition','CONTACTCLI',TOBarchitect.getValue('CONTACTCLI'));
  ExcelRangeValue(fWorkBook,'Proposition','SOCIETECLI',TOBarchitect.getValue('SOCIETECLI'));
  ExcelRangeValue(fWorkBook,'Proposition','ADRESSECLI',TOBarchitect.getValue('ADRESSECLI'));
  ExcelRangeText(fWorkBook,'Proposition','ADRESSECLI2',TOBarchitect.getValue('ADRESSECLI2'),1,1);
  ExcelRangeValue(fWorkBook,'Proposition','CPVILLECLI',TOBarchitect.getValue('CPVILLECLI'));
  ExcelRangeValue(fWorkBook,'Proposition','EMAILCLI',TOBarchitect.getValue('EMAILCLI'));
  ExcelRangeValue(fWorkBook,'Proposition','DATEPROPOSITION',TOBarchitect.getValue('DATEPROPOSITION'));
  ExcelRangeValue(fWorkBook,'Proposition','NOMREPRESENTANT',TOBarchitect.getValue('NOMREPRESENTANT'));
  ExcelRangeValue(fWorkBook,'Proposition','EMAILREPRESENTANT',TOBarchitect.getValue('EMAILREPRESENTANT'));
  ExcelRangeValue(fWorkBook,'Proposition','TYPESOFT',TOBarchitect.getValue('TYPEPROGICIEL'));
  ExcelRangeValue(fWorkBook,'Proposition','AGENCELSE',TOBarchitect.getValue('AGENCELSE'));
  ExcelRangeValue(fWorkBook,'Proposition','NUMPROPOSITION',TOBarchitect.getValue('NUMPROPOSITION'));
  ExcelRangeValue(fWorkBook,'Proposition','TELCLI',TOBarchitect.getValue('TELCLI'));
  //
  ExcelRangeValue(fWorkBook,'Part 2  BDC','NOMFAC',TOBarchitect.getValue('NOMFAC'));
  ExcelRangeValue(fWorkBook,'Part 2  BDC','ADRFAC',TOBarchitect.getValue('ADRFAC'));
  ExcelRangeValue(fWorkBook,'Part 2  BDC','CPVILLEFAC',TOBarchitect.getValue('CPVILLEFAC'));
  ExcelRangeValue(fWorkBook,'Part 2  BDC','TELFAC',TOBarchitect.getValue('TELFAC'));
  ExcelRangeValue(fWorkBook,'Part 2  BDC','EMAILFAC',TOBarchitect.getValue('EMAILFAC'));
  ExcelRangeValue(fWorkBook,'Part 2  BDC','NOMLIV',TOBarchitect.getValue('NOMLIV'));
  ExcelRangeValue(fWorkBook,'Part 2  BDC','ADRLIV',TOBarchitect.getValue('ADRLIV'));
  ExcelRangeValue(fWorkBook,'Part 2  BDC','CPVILLELIV',TOBarchitect.getValue('CPVILLELIV'));
  ExcelRangeValue(fWorkBook,'Part 2  BDC','TELLIV',TOBarchitect.getValue('TELLIV'));
  ExcelRangeValue(fWorkBook,'Part 2  BDC','EMAILLIV',TOBarchitect.getValue('EMAILLIV'));
  //
  for Indice := 0 to TOBArchitect.detail.count -1 do
  begin
  	TOBD := TOBArchitect.detail[Indice];
    if TOBD.GetValue('NOMENSEMBLE')='FRAISJOUR' then
    begin
    	TOBLL := TOBD.detail[0];
  		ExcelRangeValue(fWorkBook,'Synthese','FRAISJOUR','Forfait');
  		ExcelRangeValue(fWorkBook,'Synthese','FORFAITJOUR',TOBLL.getValue('PU'));
    end else if TOBD.GetValue('NOMENSEMBLE')='MATERIEL' then
    begin
      ExcelRangeValue(fWorkBook,'Materiels','POURCENTPREP',TOBD.getValue('POURCENTPREP'));
      ExcelRangeValue(fWorkBook,'Materiels','FRAISPREPMAT',TOBD.getValue('FRAISPREPARATION'));
      ExcelRangeValue(fWorkBook,'Materiels','REMISESMAT',TOBD.getValue('REMISES'));
      for Ind := 0 to TOBD.Detail.count -1 do
      begin
        TOBLL := TOBD.detail[Ind];
      	ExcelRangeValue(fWorkBook,'Materiels','LIBELLEMAT',TOBLL.getValue('LIBELLE'),Ind+2,1);
      	ExcelRangeValue(fWorkBook,'Materiels','QTEMAT',TOBLL.getValue('QTE'),Ind+2,1);
      	ExcelRangeValue(fWorkBook,'Materiels','PUMAT',TOBLL.getValue('PU'),Ind+2,1);
      	ExcelRangeValue(fWorkBook,'Materiels','MTMAT',TOBLL.getValue('MT'),Ind+2,1);
      end;
    end else if TOBD.GetValue('NOMENSEMBLE')='PRESTMAT' then
    begin
      for Ind := 0 to TOBD.Detail.count -1 do
      begin
        TOBLL := TOBD.detail[Ind];
      	ExcelRangeValue(fWorkBook,'Prestations materiels','LIBELLEPRESTMAT',TOBLL.getValue('LIBELLE'),Ind+2,1);
      	ExcelRangeValue(fWorkBook,'Prestations materiels','QTEPRESTMAT',TOBLL.getValue('QTE'),Ind+2,1);
      	ExcelRangeValue(fWorkBook,'Prestations materiels','PUPRESTMAT',TOBLL.getValue('PU'),Ind+2,1);
      	ExcelRangeValue(fWorkBook,'Prestations materiels','MTPRESTMAT',TOBLL.getValue('MT'),Ind+2,1);
        //
        if PosPrestations < 12 then
        begin
        	Inc(PosPrestations);
      		ExcelRangeValue(fWorkBook,'BDC PRESTA','LIBPRESTATIONS',TOBLL.getValue('LIBELLE'),PosPrestations,1);
      		ExcelRangeValue(fWorkBook,'BDC PRESTA','QTEPRESTATIONS',TOBLL.getValue('QTE'),PosPrestations,1);
      		ExcelRangeValue(fWorkBook,'BDC PRESTA','MONTANTPRESTATIONS',TOBLL.getValue('MT'),PosPrestations,1);
        end;
      end;
    end else if TOBD.GetValue('NOMENSEMBLE')='PROGICIELS' then
    begin
      ExcelRangeValue(fWorkBook,'Progiciels','REMISESPROG',TOBD.getValue('REMISES'));
      for Ind := 0 to TOBD.Detail.count -1 do
      begin
        TOBLL := TOBD.detail[Ind];
      	ExcelRangeValue(fWorkBook,'Progiciels','LIBELLESOFTS',TOBLL.getValue('LIBELLE'),Ind+2,1);
      	ExcelRangeValue(fWorkBook,'Progiciels','QTESOFTS',TOBLL.getValue('QTE'),Ind+2,1);
      	ExcelRangeValue(fWorkBook,'Progiciels','PUSOFTS',TOBLL.getValue('PU'),Ind+2,1);
      	ExcelRangeValue(fWorkBook,'Progiciels','MTSOFTS',TOBLL.getValue('MT'),Ind+2,1);
      end;
    end else if TOBD.GetValue('NOMENSEMBLE')='PRESTPROG' then
    begin
      for Ind := 0 to TOBD.Detail.count -1 do
      begin
        TOBLL := TOBD.detail[Ind];
      	ExcelRangeValue(fWorkBook,'Prestations Progiciels','LIBELLEPRESTSOFT',TOBLL.getValue('LIBELLE'),Ind+2,1);
      	ExcelRangeValue(fWorkBook,'Prestations Progiciels','QTEPRESTSOFT',TOBLL.getValue('QTE'),Ind+2,1);
      	ExcelRangeValue(fWorkBook,'Prestations Progiciels','PUPRESTSOFT',TOBLL.getValue('PU'),Ind+2,1);
      	ExcelRangeValue(fWorkBook,'Prestations Progiciels','MTPRESTSOFT',TOBLL.getValue('MT'),Ind+2,1);
        if PosPrestations < 12 then
        begin
        	Inc(PosPrestations);
      		ExcelRangeValue(fWorkBook,'BDC PRESTA','LIBPRESTATIONS',TOBLL.getValue('LIBELLE'),PosPrestations,1);
      		ExcelRangeValue(fWorkBook,'BDC PRESTA','QTEPRESTATIONS',TOBLL.getValue('QTE'),PosPrestations,1);
      		ExcelRangeValue(fWorkBook,'BDC PRESTA','MONTANTPRESTATIONS',TOBLL.getValue('MT'),PosPrestations,1);
        end;
      end;
    end else if TOBD.GetValue('NOMENSEMBLE')='ASSISTANCE' then
    begin
      for Ind := 0 to TOBD.Detail.count -1 do
      begin
        TOBLL := TOBD.detail[Ind];
      	ExcelRangeValue(fWorkBook,'Assistance','LIBELLEASSISTSOFT',TOBLL.getValue('LIBELLE'),Ind+2,1);
      	ExcelRangeValue(fWorkBook,'Assistance','QTEASSISTSOFT',TOBLL.getValue('QTE'),Ind+2,1);
      	ExcelRangeValue(fWorkBook,'Assistance','PUASSISTSOFT',TOBLL.getValue('PU'),Ind+2,1);
      	ExcelRangeValue(fWorkBook,'Assistance','MTASSISTSOFT',TOBLL.getValue('MT'),Ind+2,1);
      	ExcelRangeValue(fWorkBook,'Assistance','TYPESERVICE',TOBLL.getValue('TYPESERVICE'),Ind+2,1);
      end;
    end else if TOBD.GetValue('NOMENSEMBLE')='FORMATION' then
    begin
      for Ind := 0 to TOBD.Detail.count -1 do
      begin
        TOBLL := TOBD.detail[Ind];
      	ExcelRangeValue(fWorkBook,'Formation','LIBELLEFORMATION',TOBLL.getValue('LIBELLE'),Ind+2,1);
      	ExcelRangeValue(fWorkBook,'Formation','QTEFORMATION',TOBLL.getValue('QTE'),Ind+2,1);
      	ExcelRangeValue(fWorkBook,'Formation','LIEUFORMATION','Sur site',Ind+2,1);
      	ExcelRangeValue(fWorkBook,'Formation','NBSTAGIAIRE','1 à 5',Ind+2,1);
      	ExcelRangeValue(fWorkBook,'Formation','PUFORMATION',TOBLL.getValue('PU'),Ind+2,1);
      	ExcelRangeValue(fWorkBook,'Formation','MTFORMATION',TOBLL.getValue('MT'),Ind+2,1);
        if PosPrestations < 12 then
        begin
        	Inc(PosPrestations);
      		ExcelRangeValue(fWorkBook,'BDC PRESTA','LIBPRESTATIONS',TOBLL.getValue('LIBELLE'),PosPrestations,1);
      		ExcelRangeValue(fWorkBook,'BDC PRESTA','QTEPRESTATIONS',TOBLL.getValue('QTE'),PosPrestations,1);
      		ExcelRangeValue(fWorkBook,'BDC PRESTA','MONTANTPRESTATIONS',TOBLL.getValue('MT'),PosPrestations,1);
        end;
      end;
    end
  end;
  SelectSheet (fWorkBook,'Depart');
end;

constructor Tgenerepropal.create;
begin
  fWinExcel := unassigned;
  fModeGener := TmgConvention;
  fInternalWindow := TToolWindow97.create(Application.MainForm);
  fInternalWindow.Parent := Application.MainForm;
  fInternalWindow.Visible := false;
  fInternalWindow.Width := 600;
  //
  fTexte := THRichEditOLE.Create (fInternalWindow);
  ftexte.Parent := fInternalWindow;
  ftexte.text := '';
  ftexte.Align := alClient;
  ISBlocVide := true;
  fblocNote := THRichEditOLE.Create (fInternalWindow);
  fblocNote.Parent := fInternalWindow;
  fblocNote.text := '';
  fblocNote.Align := alClient;
end;

procedure Tgenerepropal.AddLesChampsSupArchi (TOBarchitect : TOB);
begin
  TOBarchitect.AddChampSupValeur('NUMPROPOSITION','');
  TOBarchitect.AddChampSupValeur('NUMEROCLI','');
  TOBarchitect.AddChampSupValeur('SIRETCLI','');
  TOBarchitect.AddChampSupValeur('SOCIETECLI','');
  TOBarchitect.AddChampSupValeur('ADRESSECLI','');
  TOBarchitect.AddChampSupValeur('ADRESSECLI2',' ');
  TOBarchitect.AddChampSupValeur('CPVILLECLI','');
  TOBarchitect.AddChampSupValeur('TELCLI','');
  TOBarchitect.AddChampSupValeur('EMAILCLI','');
  TOBarchitect.AddChampSupValeur('SIRENCLI','');
  TOBarchitect.AddChampSupValeur('DATEPROPOSITION','');
  TOBarchitect.AddChampSupValeur('TYPEPROGICIEL',''); // Défini si LSE BAT ou LSE BUSINESS ou les deux
  TOBarchitect.AddChampSupValeur('CONTACTCLI','');
  TOBarchitect.AddChampSupValeur('NOMFAC','');
  TOBarchitect.AddChampSupValeur('ADRFAC','');
  TOBarchitect.AddChampSupValeur('CPVILLEFAC','');
  TOBarchitect.AddChampSupValeur('TELFAC','');
  TOBarchitect.AddChampSupValeur('EMAILFAC','');
  TOBarchitect.AddChampSupValeur('NOMLIV','');
  TOBarchitect.AddChampSupValeur('ADRLIV','');
  TOBarchitect.AddChampSupValeur('CPVILLELIV','');
  TOBarchitect.AddChampSupValeur('TELLIV','');
  TOBarchitect.AddChampSupValeur('EMAILLIV','');
  TOBarchitect.AddChampSupValeur('NOMREPRESENTANT','');
  TOBarchitect.AddChampSupValeur('EMAILREPRESENTANT','');
  TOBarchitect.AddChampSupValeur('AGENCELSE','');
end;

procedure Tgenerepropal.DefiniEntetePropale (TOBPiece,TOBarchitect : TOB);
	function FormatDateTimeUs (TheDate : TdateTime) : String;
  var Day,Month,year : word;
  begin
  	DecodeDate(TheDate,Year,Month,day);
    result := IntTostr(Month)+'/'+IntTostr(Day)+'/'+IntTostr(year);
  end;

var QQ : TQuery;
		FromPerspective : boolean;
    TOBContact,TOBADRESSES,TOBADR : TOB;
    TOBRepres,TOBtiers,TOBPropal : TOB;
begin
		AddLesChampsSupArchi (TOBarchitect);
		TOBContact := TOB.create ('CONTACT',nil,-1);
  	TOBRepres := TOB.Create ('COMMERCIAL',nil,-1);
  	TOBTiers := TOB.create ('TIERS',nil,-1);
    TOBAdresses := TOB.Create ('LES ADRESSES', nil,-1);
    TOBPropal := TOB.Create ('PERSPECTIVES',nil,-1);
    TRY

      QQ := OpenSql ('SELECT T_TIERS,T_JURIDIQUE,T_LIBELLE,T_ADRESSE1,T_ADRESSE2,T_CODEPOSTAL,T_VILLE,T_SIRET FROM TIERS WHERE T_TIERS="'+
                     TOBPiece.getValue('GP_TIERS')+'" AND T_NATUREAUXI IN ("PRO","CLI")',
                     true,1,'',true);
      TOBTiers.SelectDB('',QQ);
      Ferme (QQ);
      //
      LoadLesAdresses (TOBPiece,TOBAdresses);
      //
  		if fModeGener = tMgPropal then
      begin
      	TOBarchitect.putValue('NUMPROPOSITION',TobPiece.getValue('GP_REFINTERNE'));
      end else
      begin
      	TOBarchitect.putValue('NUMPROPOSITION','Convention-'+TobPiece.getValue('GP_REFINTERNE'));
      end;
      TOBarchitect.putValue('SOCIETECLI',
      											rechdom('TTFORMEJURIDIQUE',TOBTiers.Getvalue('T_JURIDIQUE'),false)+' '+
                            TOBTiers.Getvalue('T_LIBELLE'));
      TOBarchitect.putValue('NUMEROCLI',TOBTiers.getValue('T_TIERS'));
      TOBarchitect.putValue('SIRETCLI',TOBTiers.getValue('T_SIRET'));
      TOBarchitect.putValue('ADRESSECLI',TOBTiers.getValue('T_ADRESSE1'));
      TOBarchitect.putValue('ADRESSECLI2',TOBTiers.getValue('T_ADRESSE2'));
      TOBarchitect.putValue('CPVILLECLI',TOBTiers.getValue('T_CODEPOSTAL')+
      											' '+
                            TOBTiers.getValue('T_VILLE'));
      TOBarchitect.putValue('TELCLI',TOBTiers.getValue('T_TELEPHONE'));
      TOBarchitect.putValue('EMAILCLI',TOBTiers.getValue('T_EMAIL'));
      //
      TOBarchitect.putValue('DATEPROPOSITION',FormatDateTimeUs (TobPiece.getValue('GP_DATEPIECE')));
      TOBarchitect.putValue('TYPEPROGICIEL',''); // Défini si LSE BAT ou LSE BUSINESS ou les deux
      //
      FromPerspective := false;
      if TobPiece.GetValue('GP_PERSPECTIVE') <> 0 then
      begin
        QQ := OpenSql ('SELECT * FROM PERSPECTIVES WHERE RPE_PERSPECTIVE='+
                        IntToStr(TobPiece.GetValue('GP_PERSPECTIVE')),true,-1,'',true);
        if not QQ.eof then
        begin
          FromPerspective := true;
          TOBPropal.selectDb('',QQ);
          GetContact(TOBCOntact,TobPiece.getValue('GP_TIERS'),QQ.findField('RPE_NUMEROCONTACT').Asinteger);
          TOBarchitect.putValue('CONTACTCLI',TOBCOntact.getValue('C_CIVILITE')+' '+
                                TOBCOntact.getValue('C_PRENOM')+' '+
                                TOBContact.getvalue('C_NOM'));
          TOBarchitect.putValue('TELCLI',TOBCOntact.getValue('C_TELEPHONE'));
          TOBarchitect.putValue('EMAILCLI',TOBCOntact.getValue('C_RVA'));
        end;
        Ferme (QQ);
      end;
      //
      if TOBpiece.getvalue('GP_TIERSFACTURE') <> TOBpiece.getvalue('GP_TIERS') then
      begin
        TOBADR := TOBADRESSES.detail[TOBPiece.getValue('GP_NUMADRESSEFACT')-1];
      	TOBarchitect.putValue('NOMFAC',TOBAdr.getvalue('GPA_LIBELLE'));
      	TOBarchitect.putValue('ADRFAC',TOBAdr.getvalue('GPA_ADRESSE1'));
      	TOBarchitect.putValue('CPVILLEFAC',TOBAdr.getvalue('GPA_CODEPOSTAL')+' '+TOBAdr.getvalue('GPA_VILLE'));
        if TOBADR.GetValue('GPA_NUMEROCONTACT')<> 0 then
        begin
          GetContact(TOBCOntact,TobPiece.getValue('GP_TIERSFACTURE'),TOBADR.GetValue('GPA_NUMEROCONTACT'));
      		TOBarchitect.putValue('TELFAC',TOBContact.getvalue('C_TELEPHONE'));
      		TOBarchitect.putValue('EMAILFAC',TOBContact.getvalue('C_RVA'));
        end;
      end else
      begin
      	TOBarchitect.putValue('NOMFAC',TOBarchitect.getValue('SOCIETECLI'));
      	TOBarchitect.putValue('ADRFAC',TOBarchitect.getValue('ADRESSECLI'));
      	TOBarchitect.putValue('CPVILLEFAC',TOBarchitect.getValue('CPVILLECLI'));
      	TOBarchitect.putValue('TELFAC',TOBarchitect.getValue('TELCLI'));
      	TOBarchitect.putValue('EMAILFAC',TOBarchitect.getValue('EMAILCLI'));
      end;
      //
      if TOBpiece.getvalue('GP_TIERSLIVRE') <> TOBpiece.getvalue('GP_TIERS') then
      begin
        TOBADR := TOBADRESSES.detail[TOBPiece.getValue('GP_NUMADRESSELIVR')-1];
      	TOBarchitect.putValue('NOMLIV',TOBAdr.getvalue('GPA_LIBELLE'));
      	TOBarchitect.putValue('ADRLIV',TOBAdr.getvalue('GPA_ADRESSE1'));
      	TOBarchitect.putValue('CPVILLELIV',TOBAdr.getvalue('GPA_CODEPOSTAL')+' '+TOBAdr.getvalue('GPA_VILLE'));
        if TOBADR.GetValue('GPA_NUMEROCONTACT')<> 0 then
        begin
          GetContact(TOBCOntact,TobPiece.getValue('GP_TIERSLIVRE'),TOBADR.GetValue('GPA_NUMEROCONTACT'));
      		TOBarchitect.putValue('TELLIV',TOBContact.getvalue('C_TELEPHONE'));
      		TOBarchitect.putValue('EMAILLIV',TOBContact.getvalue('C_RVA'));
        end;
      end else
      begin
      	TOBarchitect.putValue('NOMLIV',TOBarchitect.getValue('SOCIETECLI'));
      	TOBarchitect.putValue('ADRLIV',TOBarchitect.getValue('ADRESSECLI'));
      	TOBarchitect.putValue('CPVILLELIV',TOBarchitect.getValue('CPVILLECLI'));
      	TOBarchitect.putValue('TELLIV',TOBarchitect.getValue('TELCLI'));
      	TOBarchitect.putValue('EMAILLIV',TOBarchitect.getValue('EMAILCLI'));
      end;
      if not FromPerspective then
      begin
        QQ := OpenSql ('SELECT * FROM COMMERCIAL WHERE GCL_COMMERCIAL="'+
                        TobPiece.getValue('GP_REPRESENTANT')+'"',true,1,'',true);
        TOBrepres.SelectDB('',QQ);
        ferme (QQ);
        //
        if GetParamSoc('SO_GCPIECEADRESSE') then
				begin
        	GetContact(TOBCOntact,TobPiece.getValue('GP_TIERS'),TobAdresses.detail[0].getValue('GPA_NUMEROCONTACT'));
        end;
        //
        TOBarchitect.putValue('CONTACTCLI',TOBCOntact.getValue('C_CIVILITE')+' '+
                                           TOBCOntact.getValue('C_PRENOM')+' '+
                                           TOBContact.getvalue('C_NOM'));
        TOBarchitect.putValue('NOMREPRESENTANT',TOBrepres.getValue('GCL_PRENOM')+' '+
                                                TOBrepres.getValue('GCL_LIBELLE'));
        TOBarchitect.putValue('EMAILREPRESENTANT',TOBrepres.getValue('GCL_EMAIL'));
        TOBarchitect.putValue('AGENCELSE',RechDom('TTETABLISSEMENT',TOBrepres.getValue('GCL_ETABLISSEMENT'),false));
      end else
      begin
        QQ := OpenSql ('SELECT * FROM COMMERCIAL WHERE GCL_COMMERCIAL="'+
                        TOBPropal.getValue('RPE_REPRESENTANT')+'"',true,1,'',true);
        TOBrepres.SelectDB('',QQ);
        ferme (QQ);
        //
        TOBarchitect.putValue('NOMREPRESENTANT',TOBRepres.getValue('GCL_PRENOM')+' '+
                                                TOBRepres.getValue('GCL_LIBELLE'));
        TOBarchitect.putValue('EMAILREPRESENTANT',TOBRepres.getValue('GCL_EMAIL'));
        TOBarchitect.putValue('AGENCELSE',RechDom('TTETABLISSEMENT',TOBrepres.getValue('GCL_ETABLISSEMENT'),false));
      end;
		FINALLY
			FreeAndNil(fTOBContact);
  		FreeAndNil(TOBRepres);
      FreeAndNil(TOBTiers);
      FreeAndnil(TOBAdresses);
      FreeAndnil(TOBpropal);
	  end;
end;

destructor Tgenerepropal.destroy;
begin
  if not VarIsEmpty(fWinExcel) then fWinExcel.Quit;
  fWinExcel := unassigned;
  fTexte.free;
  fInternalWindow.free;
  inherited;
end;

Function Tgenerepropal.GetDocument (TOBPiece,TOBComment : TOB) : boolean;
var cledoc : r_cledoc;
		Q : TQuery;
    Req : string;
begin
  result := false;
  FillChar(CleDoc, Sizeof(CleDoc), #0);
  cledoc.NaturePiece := fnaturepiece;
  cledoc.Souche  := fSouche;
  cledoc.NumeroPiece   := fNumero;
  cledoc.Indice    := fIndiceg;
  req := 'SELECT * FROM PIECE WHERE '+WherePiece(CleDoc,ttdPiece ,False);
  Q:=OpenSQL(req,True,1,'',true) ;
  if Not Q.EOF then TOBPiece.SelectDb('',Q) else exit;
  Ferme(Q) ;
  req := 'SELECT * FROM LIGNE WHERE '+
					WherePiece(CleDoc,ttdLigne,False)+
         ' AND GL_TYPELIGNE="ART"'+
//         ' ORDER BY GL_DOMAINE,GL_FAMILLENIV1,GL_FAMILLENIV2,GL_FAMILLENIV3,GL_NUMLIGNE';
         ' ORDER BY GL_NUMLIGNE';
  Q:=OpenSQL(Req,True,-1,'',true) ;
  if Not Q.EOF then TOBPiece.LoadDetailDB('LIGNE','','',Q,false,True) else Exit;
  Ferme(Q) ;
  //
  Req := 'SELECT LO_OBJET FROM LIENSOLE WHERE LO_TABLEBLOB="GP" AND LO_QUALIFIANTBLOB="MEM" AND '+
  			 'LO_RANGBLOB=1 AND LO_IDENTIFIANT="'+
         TOBpiece.getValue('GP_NATUREPIECEG')+':'+
         TOBpiece.getValue('GP_SOUCHE')+':'+
         IntToStr(TOBpiece.getValue('GP_NUMERO'))+':'+
         IntToStr(TOBpiece.getValue('GP_INDICEG'))+'"';
  Q:=OpenSQL(req,True,1,'',true) ;
  if Not Q.EOF then TOBComment.SelectDb('',Q);
  Ferme(Q) ;
  ISBlocVide := IsBlocNoteVide(ftexte,TOBComment.getValue('LO_OBJET'));
  //
  result := true;
end;

function Tgenerepropal.IsBlocNoteVide (texte : THRichEditOLE; BlocNote : String) : boolean;
begin
	result := true;
 	StringToRich(Texte, BlocNote);
  if (Length(Texte.Text) <> 0) and (texte.Text <> #$D#$A) then result := false;
end;

procedure Tgenerepropal.EnregistreFichierExcel(TOBArchitect : TOB);
var FileName,repertLoc : string;
begin
	repertLoc := TCBPPath.GetPersonal;
	FileName := IncludeTrailingBackslash(GetParamSocSecur('SO_DIREXPORTSXLS',RepertLoc))+
  						TOBarchitect.getValue('NUMPROPOSITION')+'.xlsm';
	fWorkBook.saveas (FileName);
end;

procedure Tgenerepropal.TraiteDocument;
var TOBarchitect,TOBPiece,TOBComment : TOB;
begin
	TOBPiece := TOB.create('PIECE',nil,-1);
	TOBComment := TOB.create('LIENSOLE',nil,-1);
  //
  TOBarchitect := TOB.Create ('LE DESCRIPTIF',nil,-1);
  TRY
    if GetDocument (TOBPiece,TOBComment) then
    begin
      ConstitueStructure (TOBpiece,TOBarchitect);
      TRY
      	ConstitueFichierExcel (TOBArchitect,TOBComment);
      EXCEPT
      	PGIError('Le document n''a pas pu être créé');
        raise;
      END;
      EnregistreFichierExcel (TOBArchitect);
      PgiInfo ('Document '+TOBArchitect.getValue('NUMPROPOSITION')+'.xlsm créé avec succès');
    end;
  FINALLY
    FreeAndNil(TOBPiece);
    FreeAndNil(TOBarchitect);
  END;
end;

procedure Tgenerepropal.SetModeGener(const Value: TModegener);
begin
  fModeGener := Value;
  if fModeGener = tMgPropal then
  begin
  	fdocExcel :=  'C:\PGI00\STD\MODELEPROPALLSE.xlsm';
  end else
  begin
		fdocExcel :=  'C:\PGI00\STD\MODELECONVLSE.xlsm';
  end;

end;

{ TnewPieceGestion }

constructor TnewPieceGestion.create;
begin
	fListPiece := TlisTnewPiece.create;
  fTOBpiece := TOB.Create ('PIECE',nil,-1);
  fTOBAdresses := TOB.create ('LES ADRESSES',nil,-1);
  fTOBTiers := TOB.Create ('TIERS',nil,-1);
end;

destructor TnewPieceGestion.destroy;
begin
	fListPiece.free;
  fTOBpiece.free;
  fTOBAdresses.free;
  fTOBTiers.free;
  inherited;
end;

procedure TnewPieceGestion.ConstitueDocuments;
begin
	if fOkTraitement then
  begin
    TRY
      ChargeLeDocument;
      EclateLeDocument;
      EcritlesDocuments;
    EXCEPT
      raise;
    END;
  end;
end;

procedure TnewPieceGestion.ChargeLeDocument;
var cledoc : r_cledoc;
		Q : TQuery;
begin
  FillChar(CleDoc, Sizeof(CleDoc), #0);
  cledoc.NaturePiece := fnaturepiece;
  cledoc.Souche  := fSouche;
  cledoc.NumeroPiece   := fNumero;
  cledoc.Indice    := fIndiceg;
  LoadPieceLignes(CleDoc, fTobPiece,true,false);
  PieceAjouteSousDetail(fTOBPiece,true,false,true);
  LoadLesAdresses (fTOBPiece,fTOBAdresses);
  Q := OpenSQL('SELECT * FROM TIERS LEFT JOIN TIERSCOMPL ON T_TIERS=YTC_TIERS WHERE T_TIERS="' +
  			fTOBpiece.getValue('GP_TIERS') + '"', True,-1, '', True);
  if not Q.EOF then fTOBTiers.SelectDB('', Q);
  ferme (Q);
  fDEV.Code := fTOBPiece.getValue('GP_DEVISE');
  GetInfosDevise(fDEV);
end;

procedure TnewPieceGestion.EclateLeDocument;
var TheNewPiece : TnewPiece;
begin
	TheNewPiece := TnewPiece.create;
  TheNewPiece.TypePiece := ttypDocMat; //  matériels, logiciels externe et système + progiciels
  ThenewPiece.TOBPiece.Dupliquer(fTOBpiece,true,true);
  ThenewPiece.TOBPiece.PutValue('GP_DOMAINE','PRO');
  ThenewPiece.TOBAdresses.Dupliquer(fTOBAdresses,true,true);
  ThenewPiece.TOBTiers.Dupliquer(fTOBTiers ,true,true);
  ThenewPiece.DEV  := fDev;
  TheNewPiece.fDateAcc := fDateAcc;
  TheNewPiece.fDateDebut  := fDateDeb ;
  TheNEwPiece.fOldnaturePieceg := fTOBpiece.getvalue('GP_NATUREPIECEG');
  TheNEwPiece.fOldSouche := fTOBpiece.getvalue('GP_SOUCHE');
  TheNEwPiece.fOldNumero := fTOBpiece.getvalue('GP_NUMERO');
  TheNEwPiece.fOldIndiceG := fTOBpiece.getvalue('GP_INDICEG');
  TheNewPiece.nettoie;
  if TheNewPiece.TOBPiece.detail.count > 0 then fListPiece.Add(TheNewPiece) else TheNewPiece.free;
  //
	TheNewPiece := TnewPiece.create;
  TheNewPiece.TypePiece := TTypDocPrest; // formations et prestations
  ThenewPiece.TOBPiece.Dupliquer(fTOBpiece,true,true);
  ThenewPiece.TOBPiece.PutValue('GP_DOMAINE','PON');
  ThenewPiece.TOBAdresses.Dupliquer(fTOBAdresses,true,true);
  ThenewPiece.TOBTiers .Dupliquer(fTOBTiers ,true,true);
  ThenewPiece.DEV  := fDev;
  TheNewPiece.fDateAcc := fDateAcc;
  TheNewPiece.fDateDebut  := fDateDeb ;
  TheNEwPiece.fOldnaturePieceg := fTOBpiece.getvalue('GP_NATUREPIECEG');
  TheNEwPiece.fOldSouche := fTOBpiece.getvalue('GP_SOUCHE');
  TheNEwPiece.fOldNumero := fTOBpiece.getvalue('GP_NUMERO');
  TheNEwPiece.fOldIndiceG := fTOBpiece.getvalue('GP_INDICEG');
  TheNewPiece.nettoie;
  if TheNewPiece.TOBPiece.detail.count > 0 then fListPiece.Add(TheNewPiece) else TheNewPiece.free;
  //
	TheNewPiece := TnewPiece.create;
  TheNewPiece.TypePiece := TTypDocCont; // contrats de maintenance
  ThenewPiece.TOBPiece.PutValue('GP_DOMAINE','REC');
  ThenewPiece.TOBPiece.Dupliquer(fTOBpiece,true,true);
  ThenewPiece.TOBAdresses.Dupliquer(fTOBAdresses,true,true);
  ThenewPiece.TOBTiers .Dupliquer(fTOBTiers ,true,true);
  ThenewPiece.DEV  := fDev;
  TheNewPiece.fDateAcc := fDateAcc;
  TheNewPiece.fDateDebut  := fDateDeb ;
  TheNEwPiece.fOldnaturePieceg := fTOBpiece.getvalue('GP_NATUREPIECEG');
  TheNEwPiece.fOldSouche := fTOBpiece.getvalue('GP_SOUCHE');
  TheNEwPiece.fOldNumero := fTOBpiece.getvalue('GP_NUMERO');
  TheNEwPiece.fOldIndiceG := fTOBpiece.getvalue('GP_INDICEG');
  TheNewPiece.nettoie;
  if TheNewPiece.TOBPiece.detail.count > 0 then fListPiece.Add(TheNewPiece) else TheNewPiece.free;
end;

procedure TnewPieceGestion.EcritlesDocuments;
begin
  //
  BEGINTRANS;
	TRY
		BeforeEcriture;
    EcritureDocuments;
    PositionneEtudeEnTermine;
  	COMMITTRANS;
  EXCEPT
  	ROLLBACK;
    raise
  END;
end;


procedure TnewPieceGestion.BeforeEcriture;
var Indice : integer;
		ThePiece : TnewPiece;
begin
	TRY
    for Indice := 0 to FListPiece.Count -1 do
    begin
      ThePiece := fListPiece.Items [Indice];
      if ThePiece.TOBpiece.detail.count = 0 then continue; // pas la peine de continue le traitement pas de ligne
      if not AFMajProspectClient(ThePiece.TOBPiece.GetString('GP_TIERS')) then Continue;
      ThePiece.AffecteNaturePiece;
      ThePiece.AffecteEtatPiece('');
      ThePiece.SetDatePiece(fDateAcc);
      if ThePiece.fType = TTypDocCont then
      begin
  			ThePiece.ConstitueCodeContrat (fAffaireContrat);
      	ThePiece.SetAffaire(FaffaireContrat);
      end else
      begin
      	ThePiece.SetAffaire(fAffaire);
      end;
      ThePiece.recalcule;
      ThePiece.SetInfosPiece;
      if ThePiece.fType = TTypDocCont then
      begin
      	ThePiece.ConstitueContrat (FaffaireContrat,Faffaire);
        ThePiece.TOBPiece.putValue('GP_AFFAIREDEVIS','');
      end else
      begin
				ThePiece.DefiniAffairePiece;
      end;
      ThePiece.InitToutModif;
    end;
  EXCEPT
	  on E: Exception do
    begin
    	PgiError(E.Message);
			raise;
    end;
  END;
end;

procedure TnewPieceGestion.EcritureDocuments;
var Indice : integer;
		ThePiece : TnewPiece;
begin
	TRY
    for Indice := 0 to FListPiece.Count -1 do
    begin
      ThePiece := fListPiece.Items [Indice];
      if ThePiece.tobPiece.detail.count = 0 then continue;
      ThePiece.EcritDocuments;
    end;
  EXCEPT
  	raise;
  END;
end;



procedure TnewPieceGestion.DemandeInfosSupl;
var TOBinfos : TOB;
		QQ : Tquery;
    Affaire,Tiers,Libelle,CodeAffaire2 : string;
begin
	fOkTraitement := false;
  QQ := OpenSql ('SELECT GP_AFFAIRE,GP_TIERS,T_LIBELLE,ET_CHARLIBRE1 FROM PIECE '+
  							 'LEFT JOIN TIERS ON T_TIERS=GP_TIERS AND T_NATUREAUXI="CLI" '+
                 'LEFT JOIN ETABLISS ON ET_ETABLISSEMENT=GP_ETABLISSEMENT '+
                 'WHERE GP_NATUREPIECEG="'+fnaturepiece+'" AND '+
  							 'GP_SOUCHE="'+fSouche+'" AND GP_NUMERO='+InttoStr(fNumero)+' AND '+
                 'GP_INDICEG='+InttoStr(fIndiceg),true,1,'',true);
  if not QQ.eof then
  begin
  	Affaire:= QQ.findField('GP_AFFAIRE').AsString;
  	Tiers:= QQ.findField('GP_TIERS').AsString;
  	Libelle:= QQ.findField('T_LIBELLE').AsString;
  	CodeAffaire2:= QQ.findField('ET_CHARLIBRE1').AsString;
  end;
  ferme (QQ);
  TOBInfos := TOB.Create ('LES PARAMS',nil,-1);
  TOBInfos.AddChampSupValeur('DATE',now);
  TOBInfos.AddChampSupValeur('DATECON',now);
  TOBInfos.AddChampSupValeur('NUMETUDE',fNumero);
  TOBInfos.AddChampSupValeur('AFFAIRE',Affaire);
  TOBInfos.AddChampSupValeur('CODEAFFAIRE2',CodeAffaire2);
  TOBInfos.AddChampSupValeur('TIERS',Tiers);
  TOBInfos.AddChampSupValeur('LIBELLETIERS',Libelle);
  TOBInfos.AddChampSupValeur('OK','-');
  TheTOB := TOBinfos;
  AGLLanceFiche('BTP','BTACCETU_COMPL','','','ACTION=MODIFICATION');
  TheTOB := nil;
  if TOBinfos.getValue('OK')='X' then
  begin
    fAffaire := TOBInfos.getvalue('AFFAIRE');
    fDateAcc := TOBInfos.getvalue('DATE');
    fDateDeb := TOBInfos.getvalue('DATECON');
    //
    fOkTraitement := true;
  end;
  freeAndNil(TOBinfos);
end;

procedure TnewPieceGestion.PositionneEtudeEnTermine;
var SQl : String;
begin
	Sql := 'UPDATE AFFAIRE SET AFF_ETATAFFAIRE="TER" WHERE AFF_AFFAIRE="'+fTOBPiece.getValue('GP_AFFAIREDEVIS')+'"';
	ExecuteSQL(Sql);
	Sql := 'UPDATE PIECE SET GP_VIVANTE="-" WHERE GP_NATUREPIECEG="'+fnaturepiece+'" '+
  			 'AND GP_SOUCHE="'+fSouche+'" AND GP_NUMERO='+inttostr(fNumero)+' '+
         'AND GP_INDICEG='+IntTOstr(fIndiceg);
	ExecuteSQL(Sql);
	Sql := 'UPDATE LIGNE SET GL_VIVANTE="-" WHERE GL_NATUREPIECEG="'+fnaturepiece+'" '+
  			 'AND GL_SOUCHE="'+fSouche+'" AND GL_NUMERO='+inttostr(fNumero)+' '+
         'AND GL_INDICEG='+IntTOstr(fIndiceg);
	ExecuteSQL(Sql);
end;

{ TnewPiece }

constructor TnewPiece.create;
begin
	fnewNum := 0;
  fTOBAffaireAdrInt := TOB.create ('ADRESSES',nil,-1);
	fTOBAffaireP := TOB.Create ('AFFAIRE',nil,-1);
  fTOBaffaire := TOB.create ('AFFAIRE',nil,-1);
  fTOBFACTAFF := TOB.Create ('LES ECHEANCES',nil,-1);

	fTOBPiece := TOB.create ('PIECE',nil,-1);
	fTOBAdresses := TOB.Create ('LES ADRESSES',nil,-1);
  fTOBOuvrages := TOB.Create ('LES OUVRAGES',nil,-1);
  fTOBOuvragesP := TOB.Create ('LES OUVRAGES PLAT',nil,-1);
  fTOBBases := TOB.Create ('LES BASE DE TAXES',nil,-1);
  fTOBBasesL := TOB.Create ('LES BASES TAXES LIG',nil,-1);
  fTOBTiers := TOB.Create ('TIERS',nil,-1);
  fTOBArticles := TOB.Create ('LES ARTICLES',nil,-1);
  fTOBPorcs := TOB.Create ('LES PORCS',nil,-1);
  fTOBPieceRG := TOB.Create ('LES PIECE RG',nil,-1);
  fTOBBasesRG := TOB.Create ('LES BASES RG',nil,-1);
  fTOBEches :=TOB.Create ('LES ECHEANCES',nil,-1);
  fTOBPieceTrait := TOB.Create ('LES PIECETRAIT',nil,-1);
  fTOBSSTrait := TOB.Create('LES SS TRAITS',nil,-1);
  fTOBVTECOLL := TOB.Create('LES VTE COLL',nil,-1);

end;

destructor TnewPiece.destroy;
begin
	freeAndNil(fTOBAffaireAdrInt);
	FreeAndNil(fTOBAffaireP);
  FreeAndNil(fTOBFACTAFF);
	FreeAndNil(fTOBAffaire);
	FreeAndNil(fTOBPiece);
	FreeAndNil(fTOBAdresses);
  FreeAndNil(fTOBOuvrages);
  FreeAndNil(fTOBOuvragesP);
  FreeAndNil(fTOBBases);
  FreeAndNil(fTOBBasesL);
  FreeAndNil(fTOBTiers);
  FreeAndNil(fTOBArticles);
  FreeAndNil(fTOBPorcs);
  FreeAndNil(fTOBPieceRG);
  FreeAndNil(fTOBBasesRG);
  FreeAndNil(fTOBEches);
  freeAndNil(fTOBPieceTrait);
  FreeAndNil(fTOBSSTrait);
  fTOBVTECOLL.Free;
  inherited;
end;


procedure TnewPiece.nettoie;
var Indice,iTypeL,iDomaine,iFamille1,iFamille2,INbLig : integer;
		TOBL : TOB;
    TypeLigne,Domaine,Famille1,Famille2 : string;
begin
	Indice := 0;
  iTypeL := -1;
  iDomaine := -1;
  iFamille1 := -1;
  iFamille2 := -1;
	repeat
  	TOBL := fTOBPiece.detail[Indice];
    if Indice = 0 then
    begin
    	iTypeL := TOBL.GetNumChamp('GL_TYPELIGNE');
    	iDomaine := TOBL.GetNumChamp('GL_DOMAINE');
    	iFamille1 := TOBL.GetNumChamp('GL_FAMILLENIV1');
    	iFamille2 := TOBL.GetNumChamp('GL_FAMILLENIV2');
    end;
    //
    TypeLigne := TOBL.GetValeur(itypeL);
    Domaine := TOBL.GetValeur(iDomaine);
    Famille1 := TOBL.GetValeur(iFamille1);
    Famille2 := TOBL.GetValeur(iFamille2);
    //
    if TypeLigne <> 'ART' then begin inc(Indice); continue; end;
    //
    if TypeLigne = 'ART' then
    begin
      if fType= ttypDocMat then
      begin
      	if (Domaine='PRO') then
        begin
        	Inc(Indice);
        	continue;
        end;
        TOBL.free;
      end else if fType = TTypDocPrest then
      begin
      	if (Domaine='PON') then
        begin
        	Inc(Indice);
        	continue;
        end;
        TOBL.free;
      end else if fType = TTypDocCont then
      begin
      	if (Domaine='REC') then
        begin
        	Inc(Indice);
        	continue;
        end;
        TOBL.free;
      end;
    end;
  until Indice >= fTOBPiece.detail.count;
  //
  NettoieParagraphe (fTOBpiece);
  //
  INbLig := 0;
  for Indice := 0 to fTOBpiece.detail.count -1 do
  begin
  	TOBL := fTOBPiece.detail[Indice];
    if Indice = 0 then
    begin
    	iTypeL := TOBL.GetNumChamp('GL_TYPELIGNE');
    end;
    TypeLigne := TOBL.GetValeur(itypeL);
    if TypeLigne = 'ART' then inc(Inblig);
  end;
  if INBLig = 0 then fTOBPiece.clearDetail;
end;

function TnewPiece.Controleparagraphe(TOBP : TOB;Debut,Niveau : integer; Var FinParagraphe : integer) : boolean;
var Indice,NivSuivant,FinSUiv : integer;
		TOBL : TOB;
begin
	result := false;
  Indice := Debut;
  repeat
  	TOBL := TOBP.detail[Indice];
  	if IsFinParagraphe (TOBL,Niveau) then
    BEGIN
    	FinParagraphe := Indice;
    	break;
    END;
    if IsDebutParagraphe (TOBL,Niveau) then BEGIN Inc(Indice); continue; END;
    if IsDebutParagraphe (TOBL) then
    begin
      NivSuivant := TOBL.GetValue('GL_NIVEAUIMBRIC');
      if Controleparagraphe ( TOBP,Indice,NivSuivant,FinSuiv) then
      begin
        // il y a bien qq chose a traiter
      	result := true;
        Indice := FinSuiv;
      end else
      begin
        Indice := FinSuiv;
      end;
    end;
    if TOBL.getValue('GL_NIVEAUIMBRIC') > Niveau then BEGIN Inc(Indice); continue; END;
    if TOBL.GetValue('GL_TYPELIGNE') <> 'ART' Then BEGIN Inc(Indice); Continue; END; // on saute les lignes de commentaires
    result := true;
    inc(Indice);
  until Indice >= TOBP.detail.count;
end;

procedure TnewPiece.DeleteParagraphe (TOBP : TOB;Debut,Niveau : integer);
var Indice : integer;
		TOBl : TOB;
    StopIt : boolean;
begin
	Indice := Debut;
  StopIt := false;
	Repeat
  	TOBL := TOBP.detail[Indice];
    if IsFinParagraphe (TOBL,Niveau) then
    begin
    	StopIt := true;
    end;
    TOBL.free;
  until (Indice >= TOBP.detail.count) or (StopIt);
end;

procedure TnewPiece.NettoieParagraphe(TOBP : TOB);
var indice,FinParag : integer;
		TOBL : TOB;
begin
	Indice := 0;
  if TOBP.detail.count = 0 then exit;
  repeat
    TOBL := TOBP.detail[Indice];
    if TOBL.getValue('GL_NIVEAUIMBRIC') = 1 then
    begin
      if IsDebutParagraphe (TOBL,1) then
      begin
        if not Controleparagraphe(TOBP,Indice,1,FinParag) then
        begin
          DeleteParagraphe (TOBP,Indice,1);
        end else
        begin
        	Indice := FinParag; // saute a la fin du paragraphe traité
        end;
      end else
      begin
        inc(Indice);
      end;
    end else inc(Indice);
  until Indice >= TOBP.detail.count;
end;


procedure TnewPiece.ReCalcule;
var TOBL : TOB;
		indice : integer;
begin
  NumeroteLignesGC(nil,fTOBpiece,true, false);
	CalculeMontantsDoc (fTOBpiece,fTOBOuvrages,false);
  ZeroFacture (fTOBpiece);
  for Indice := 0 to fTOBpiece.detail.count -1 do
	begin
  	TOBL := fTOBPiece.detail[Indice];
  	ZeroLigneMontant (TOBL);  // reinit des montants de la ligne
  end;
  ZeroMontantPorts (fTOBPorcs);
  PutValueDetail(fTOBPiece, 'GP_RECALCULER', 'X');
  fTOBPieceTrait.ClearDetail;
  fTOBSSTrait.ClearDetail;
  fTOBVTECOLL.ClearDetail;
	CalculFacture (nil,fTOBPiece,fTOBPieceTrait,fTOBSSTrait,fTOBouvrages,fTOBouvragesP,
  							 fTOBBases,fTOBBasesL,fTOBTiers,fTOBArticles,
                 fTOBPorcs,fTOBPieceRG,fTOBBasesRG, fTOBVTECOLL, fDEV);
	GereEcheancesGC(fTOBpiece,fTOBTiers,fTOBEches,nil,fTOBPieceRG,fTOBPieceTrait,nil,taCreat,DEV,False) ;

  CalculeSousTotauxPiece(fTOBPiece);
  RecalculeRG(fTOBPORCS,fTOBPIECE, fTOBPIECERG, fTOBBASES, fTOBBASESRG,nil, fDEV);
end;

procedure TnewPiece.AffecteNaturePiece;
begin
	if (fType = ttypDocMat) or (ftype = TTypDocPrest)  then
  begin
  	fTOBPiece.PutValue('GP_NATUREPIECEG','DBT'); // Devis standard
  end else if (fType = TTypDocCont) then
  begin
  	fTOBPiece.PutValue('GP_NATUREPIECEG','AFF');  // Devis de contrat
  end;
  fTOBPiece.PutValue('GP_SOUCHE',GetSoucheG(fTOBPiece.GetValue('GP_NATUREPIECEG'),
                                 fTOBPiece.GetValue('GP_ETABLISSEMENT'),
                                 fTOBPiece.GetValue('GP_DOMAINE')));
  fTOBPiece.PutValue('GP_NUMERO',0);
  fTOBPiece.PutValue('GP_INDICEG',0);

end;

procedure TnewPiece.AffecteEtatPiece(CodeEtat : string);
begin
  fTOBPiece.PutValue('GP_ETATVISA',CodeEtat);
end;

procedure TnewPiece.SetInfosPiece;
begin
	fNewNum:=GetNumSoucheG(fTOBPiece.GetValue('GP_SOUCHE'),fTOBPiece.GetValue('GP_DATEPIECE')) ;
  fTOBPiece.putValue('GP_NUMERO',fNewNum);
  IncNumSoucheG (fTOBPiece.GetValue('GP_SOUCHE'),fTOBPiece.GetValue('GP_DATEPIECE'));
  SetDefinitiveNumberAndOthers;
end;

procedure TnewPiece.InitToutModif ;
BEGIN
  fTOBpiece.SetAllModifie(True) ;
  fTOBBases.SetAllModifie(True)  ;
  fTOBBasesL.SetAllModifie(True)  ;
  fTOBOuvragesP.SetAllModifie(True)  ;
  fTOBEches.SetAllModifie(True)  ;
  fTOBPorcs.SetAllModifie(True)  ;
  // Modif BTP
  fTOBPieceRG.SetAllModifie (true);
  fTOBBasesRG.SetAllModifie (true);
  fTOBAffaire.SetAllModifie (true);
  fTOBFACTAFF.SetAllModifie(true);
  fTOBAffaireP.SetAllModifie (true);
  fTOBAdresses.setAllModifie(True);
  fTOBOuvrages.setAllModifie(True);
END ;


procedure TnewPiece.SetDefinitiveNumberAndOthers;

	procedure TraiteDetailOuvrage (TOBD : TOB; Num : integer; Nature,Souche,Affaire,Aff1,Aff2,Aff3,Avenant : string; DatePiece : TdateTime);
  var i, iNum,iNat,Isouche,Iaffaire,Iaff1,Iaff2,Iaff3,Iavenant,IdatePiece: integer;
  		TOBB : TOB;
  begin
  	INum := 0;
  	INat := 0;
  	ISouche := 0;
    Iaffaire := 0;
    Iaff1  := 0;
    Iaff2  := 0;
    Iaff3  := 0;
    Iavenant  := 0;
    IdatePiece  := 0;
  	for I := 0 to TOBD.detail.count -1 do
    begin
      if i = 0 then
      begin
        iNum := TOBD.Detail[i].GetNumChamp('BLO_NUMERO');
        iNat := TOBD.Detail[i].GetNumChamp('BLO_NATUREPIECEG');
        iSouche := TOBD.Detail[i].GetNumChamp('BLO_SOUCHE');
        iAffaire := TOBD.Detail[i].GetNumChamp('BLO_AFFAIRE');
        iAff1 := TOBD.Detail[i].GetNumChamp('BLO_AFFAIRE1');
        iAff2 := TOBD.Detail[i].GetNumChamp('BLO_AFFAIRE2');
        iAff3 := TOBD.Detail[i].GetNumChamp('BLO_AFFAIRE3');
        iAvenant := TOBD.Detail[i].GetNumChamp('BLO_AVENANT');
        iDatePiece := TOBD.Detail[i].GetNumChamp('BLO_DATEPIECE');
      end;
      TOBB := TOBD.detail[I];
      TOBB.PutValeur(iNum, Num);
      TOBB.PutValeur(iNat, nature);
      TOBB.PutValeur(iSouche, Souche);
      TOBB.PutValeur(iAffaire, Affaire);
      TOBB.PutValeur(iAff1, Aff1);
      TOBB.PutValeur(iAff2, Aff2);
      TOBB.PutValeur(iAff3, Aff3);
      TOBB.PutValeur(iAvenant, Avenant);
      TOBB.PutValeur(iDatePiece, DatePiece);
      if TOBB.detail.count > 0 then TraiteDetailOuvrage(TOBB,Num,nature,Souche,Affaire,aff1,aff2,aff3,avenant,datepiece);
      if TOBD.fieldExists('UTILISE') then TOBD.PutValue('UTILISE','X');
    end;
  end;

	procedure TraiteDetailOuvragePlat (TOBD : TOB; Num : integer; Nature,Souche,Affaire,Aff1,Aff2,Aff3,Avenant : string; DatePiece : TdateTime);
  var i, iNum,iNat,Isouche,Iaffaire,Iaff1,Iaff2,Iaff3,Iavenant,IdatePiece: integer;
  		TOBB : TOB;
  begin
  	INum := 0;
  	INat := 0;
  	ISouche := 0;
    Iaffaire := 0;
    Iaff1  := 0;
    Iaff2  := 0;
    Iaff3  := 0;
    Iavenant  := 0;
    IdatePiece  := 0;

  	for I := 0 to TOBD.detail.count -1 do
    begin
      if i = 0 then
      begin
        iNum := TOBD.Detail[i].GetNumChamp('BOP_NUMERO');
        iNat := TOBD.Detail[i].GetNumChamp('BOP_NATUREPIECEG');
        iSouche := TOBD.Detail[i].GetNumChamp('BOP_SOUCHE');
        iAffaire := TOBD.Detail[i].GetNumChamp('BOP_AFFAIRE');
        iAff1 := TOBD.Detail[i].GetNumChamp('BOP_AFFAIRE1');
        iAff2 := TOBD.Detail[i].GetNumChamp('BOP_AFFAIRE2');
        iAff3 := TOBD.Detail[i].GetNumChamp('BOP_AFFAIRE3');
        iAvenant := TOBD.Detail[i].GetNumChamp('BOP_AVENANT');
        iDatePiece := TOBD.Detail[i].GetNumChamp('BOP_DATEPIECE');
      end;
      TOBB := TOBD.detail[I];
      TOBB.PutValeur(iNum, Num);
      TOBB.PutValeur(iNat, nature);
      TOBB.PutValeur(iSouche, Souche);
      TOBB.PutValeur(iAffaire, Affaire);
      TOBB.PutValeur(iAff1, Aff1);
      TOBB.PutValeur(iAff2, Aff2);
      TOBB.PutValeur(iAff3, Aff3);
      TOBB.PutValeur(iAvenant, Avenant);
      TOBB.PutValeur(iDatePiece, DatePiece);
    end;
  end;

  procedure  TraiteBasesL (TOBD : TOB; Num : integer; Nature,Souche : string; DatePiece : TDateTime);
  var i, iNum,iNat,Isouche,IdatePiece: integer;
  		TOBB : TOB;
  begin
  	INum := 0;
  	INat := 0;
  	ISouche := 0;
    IDatePiece := 0;
  	for I := 0 to TOBD.detail.count -1 do
    begin
      if i = 0 then
      begin
        iNum := TOBD.Detail[i].GetNumChamp('BLB_NUMERO');
        iNat := TOBD.Detail[i].GetNumChamp('BLB_NATUREPIECEG');
        iSouche := TOBD.Detail[i].GetNumChamp('BLB_SOUCHE');
        iDatePiece := TOBD.Detail[i].GetNumChamp('BLB_DATEPIECE');
      end;
      TOBB := TOBD.detail[I];
      TOBB.PutValeur(iNum, Num);
      TOBB.PutValeur(iNat, nature);
      TOBB.PutValeur(iSouche, Souche);
      TOBB.PutValeur(iDatePiece, DatePiece);
    end;
  end;

var SoucheG: String3;
  i, iNum,iNat,Isouche,INumNomen,NumDef: integer;
  DatePiece : Tdatetime;
  iAffaire,Iaffaire1,Iaffaire2,Iaffaire3,IAvenant,IdatePiece : integer;
  TOBB,TOBL: TOB;
  nature,Affaire,Aff1,Aff2,Aff3,Avenant : string;
begin
  iNum := 0; Inat := 0; ISouche := 0; INumNomen := 0;
  Iaffaire := 0;
  Iaffaire1  := 0;
  Iaffaire2  := 0;
  Iaffaire3  := 0;
  Iavenant  := 0;
  IdatePiece  := 0;

  SoucheG := fTOBPiece.GetValue('GP_SOUCHE');
  if SoucheG = '' then Exit;
  Nature := fTOBPiece.GetValue('GP_NATUREPIECEG');
  NumDef := fTOBPiece.GetValue('GP_NUMERO');
  Affaire := fTOBPiece.GetValue('GP_AFFAIRE');
  Aff1 := fTOBPiece.GetValue('GP_AFFAIRE1');
  Aff2 := fTOBPiece.GetValue('GP_AFFAIRE2');
  Aff3 := fTOBPiece.GetValue('GP_AFFAIRE3');
  Avenant := fTOBPiece.GetValue('GP_AVENANT');
  DatePiece := fTOBPiece.GetValue('GP_DATEPIECE');
  //
  fTOBPiece.PutValue('GP_NUMERO', NumDef);
  for i := 0 to fTOBPiece.Detail.Count - 1 do
  begin
    if i = 0 then
    begin
      iNum := fTOBPiece.Detail[i].GetNumChamp('GL_NUMERO');
      iNat := fTOBPiece.Detail[i].GetNumChamp('GL_NATUREPIECEG');
      iSouche := fTOBPiece.Detail[i].GetNumChamp('GL_SOUCHE');
      iNumNomen := fTOBPiece.Detail[i].GetNumChamp('GL_INDICENOMEN');
      //
      iAffaire := fTOBPiece.Detail[i].GetNumChamp('GL_AFFAIRE');
      iAffaire1 := fTOBPiece.Detail[i].GetNumChamp('GL_AFFAIRE1');
      iAffaire2 := fTOBPiece.Detail[i].GetNumChamp('GL_AFFAIRE2');
      iAffaire3 := fTOBPiece.Detail[i].GetNumChamp('GL_AFFAIRE3');
      iAvenant := fTOBPiece.Detail[i].GetNumChamp('GL_AVENANT');
      iDatePiece := fTOBPiece.Detail[i].GetNumChamp('GL_DATEPIECE');
    end;
    //
    TOBL := fTOBPiece.Detail[i];
    TOBL.PutValeur(iNum, NumDef);
    TOBL.PutValeur(iNat, nature);
    TOBL.PutValeur(iSouche, SoucheG);
    TOBL.PutValeur(iAffaire, Affaire);
    TOBL.PutValeur(iAffaire1, Aff1);
    TOBL.PutValeur(iAffaire2, Aff2);
    TOBL.PutValeur(iAffaire3, Aff3);
    TOBL.PutValeur(iAvenant, Avenant);
    TOBL.PutValeur(iDatePiece, DatePiece);
    //
    if (IsOuvrage (TOBL)) and (fTOBOuvrages<>nil) and (fTOBOuvrages.detail.count > 0) then
    begin
    	if TOBL.getValeur(INumNomen) < fTOBOuvrages.detail.count then
      begin
      	TOBB := fTOBOuvrages.Detail[TOBL.getValeur(INumNomen)+1];
        TraiteDetailOuvrage (TOBB,NumDef,Nature,SoucheG,Affaire,Aff1,Aff2,Aff3,Avenant,datepiece);
      end;
      if (fTOBOuvragesP <> nil) and (fTOBouvragesP.detail.count > 0) then
      begin
      	TOBB := fTOBOuvragesP.findFirst(['NUMORDRE'],[TOBL.GetValue('GL_NUMORDRE')],true);
        if TOBB <> nil then
        begin
        	TraiteDetailOuvragePlat (TOBB,NumDef,Nature,SoucheG,Affaire,Aff1,Aff2,Aff3,Avenant,datepiece);
        end;
      end;
    end;
  end;
  //
  if fTOBAdresses  <> nil then
  begin
    for i := 0 to fTOBAdresses.Detail.Count - 1 do
    begin
      if i = 0 then
      begin
        iNum := fTOBAdresses.Detail[i].GetNumChamp('GPA_NUMERO');
        iNat := fTOBAdresses.Detail[i].GetNumChamp('GPA_NATUREPIECEG');
        iSouche := fTOBAdresses.Detail[i].GetNumChamp('GPA_SOUCHE');
//        iDatePiece := fTOBBases.Detail[i].GetNumChamp('GPA_DATEPIECE');
      end;
      TOBB := fTOBAdresses.Detail[i];
      TOBB.PutValeur(iNum, NumDef);
      TOBB.PutValeur(iNat, nature);
      TOBB.PutValeur(iSouche, SoucheG);
//      TOBB.PutValeur(IdatePiece, DatePiece);
    end;
  end;
  //
  if fTOBBases <> nil then
  begin
    for i := 0 to fTOBBases.Detail.Count - 1 do
    begin
      if i = 0 then
      begin
        iNum := fTOBBases.Detail[i].GetNumChamp('GPB_NUMERO');
        iNat := fTOBBases.Detail[i].GetNumChamp('GPB_NATUREPIECEG');
        iSouche := fTOBBases.Detail[i].GetNumChamp('GPB_SOUCHE');
        iDatePiece := fTOBBases.Detail[i].GetNumChamp('GPB_DATEPIECE');
      end;
      TOBB := fTOBBases.Detail[i];
      TOBB.PutValeur(iNum, NumDef);
      TOBB.PutValeur(iNat, nature);
      TOBB.PutValeur(iSouche, SoucheG);
      TOBB.PutValeur(IdatePiece, DatePiece);
    end;
  end;
  if fTOBBasesL <> nil then
  begin
    for i := 0 to fTOBBasesL.Detail.Count - 1 do
    begin
      TOBB := fTOBBasesL.Detail[i];
      TraiteBasesL (TOBB,NumDef,Nature,SoucheG,DatePiece);
    end;
  end;
  if fTOBEches <> nil then
  begin
    for i := 0 to fTOBEches.Detail.Count - 1 do
    begin
      if i = 0 then
      begin
        iNum := fTOBEches.Detail[i].GetNumChamp('GPE_NUMERO');
        iNat := fTOBEches.Detail[i].GetNumChamp('GPE_NATUREPIECEG');
        iSouche := fTOBEches.Detail[i].GetNumChamp('GPE_SOUCHE');
        iDatePiece := fTOBEches.Detail[i].GetNumChamp('GPE_DATEPIECE');
      end;
      TOBB := fTOBEches.Detail[i];
      TOBB.PutValeur(iNum, NumDef);
      TOBB.PutValeur(iNat, nature);
      TOBB.PutValeur(iSouche, SoucheG);
      TOBB.PutValeur(iDatePiece, DatePiece);
    end;
  end;
  if fTOBPieceRG <> nil then
  begin
    for i := 0 to fTOBPieceRG.Detail.Count - 1 do
    begin
      if i = 0 then
      begin
        iNum := fTOBPIECERg.Detail[i].GetNumChamp('PRG_NUMERO');
        iNat := fTOBPIECERg.Detail[i].GetNumChamp('PRG_NATUREPIECEG');
        iSouche := fTOBPIECERg.Detail[i].GetNumChamp('PRG_SOUCHE');
        iDatePiece := fTOBPIECERg.Detail[i].GetNumChamp('PRG_DATEPIECE');
      end;
      TOBB := fTOBPIECERG.Detail[i];
      TOBB.PutValeur(iNum, NumDef);
      TOBB.PutValeur(iNat, nature);
      TOBB.PutValeur(iSouche, SoucheG);
      TOBB.PutValeur(iDatePiece, DatePiece);
    end;
  end;
  if fTOBBasesRg <> nil then
  begin
    for i := 0 to fTOBBasesRg.Detail.Count - 1 do
    begin
      if i = 0 then
      begin
        iNum := fTOBBasesRg.Detail[i].GetNumChamp('PBR_NUMERO');
        iNat := fTOBBasesRg.Detail[i].GetNumChamp('PBR_NATUREPIECEG');
        iSouche := fTOBBasesRg.Detail[i].GetNumChamp('PBR_SOUCHE');
        iDatePiece := fTOBBasesRg.Detail[i].GetNumChamp('PBR_DATEPIECE');
      end;
      TOBB := fTOBBasesRG.Detail[i];
      TOBB.PutValeur(iNum, NumDef);
      TOBB.PutValeur(iNat, nature);
      TOBB.PutValeur(iSouche, SoucheG);
      TOBB.PutValeur(iDatePiece, DatePiece);
    end;
  end;
end;

procedure TnewPiece.EcritDocuments;
begin
	V_PGI.IOError := OeOk;
	if fType = TTypDocCont then
  begin
		// Ecriture de l'affaire contrat et de ses échéances
    if (not fTOBAffaire.InsertDB(nil)) then V_PGI.ioerror := OeUnknown;
    if (V_PGI.ioerror = OeOk) and (fTOBFACTAFF.detail.count > 0) then
    begin
    	if (V_PGI.ioerror = OeOk) and (not fTOBFACTAFF.InsertDB(nil)) then V_PGI.ioerror := OeUnknown;
    end;
  end else
  begin
    if (not fTOBAffaireP.InsertDB(nil)) then V_PGI.ioerror := OeUnknown;
  end;
  if (V_PGI.ioerror = OeOk) and (not fTOBPiece.InsertDBByNivel(false)) then V_PGI.IOError := OeUnknown;
  if (V_PGI.ioerror = OeOk) and (fTOBAdresses.detail.count > 0) then
  begin
  	if fTOBAdresses.detail.count > 1 then
    begin
    	if fTOBAdresses.detail[0].getValue('GPA_TYPEPIECEADR')=fTOBAdresses.detail[0].getValue('GPA_TYPEPIECEADR') then
      begin
      	fTOBAdresses.detail[1].free;
      end;
    end;
  	if not fTOBAdresses.insertDb(nil) then V_PGI.IOError := OeUnknown;
  end;
  if (V_PGI.ioerror = OeOk) and (fTOBOuvrages.detail.count > 0) then if not fTOBOuvrages.InsertDB(nil) then V_PGI.IOError := OeUnknown;
  if (V_PGI.ioerror = OeOk) and (fTOBOuvragesP.detail.count > 0) then if not fTOBOuvragesP.InsertDB(nil) then V_PGI.IOError := OeUnknown;
  if (V_PGI.ioerror = OeOk) and (fTOBBases.detail.count > 0) then if not fTOBBases.InsertDB(nil) then V_PGI.IOError := OeUnknown;
  if (V_PGI.ioerror = OeOk) and (fTOBBasesL.detail.count > 0) then if not fTOBBasesL.InsertDB(nil) then V_PGI.IOError := OeUnknown;
  if (V_PGI.ioerror = OeOk) and (fTOBEches.detail.count > 0) then if not fTOBEches.InsertDB(nil) then V_PGI.IOError := OeUnknown;
  if (V_PGI.ioerror = OeOk) and (fTOBPieceRG.detail.count > 0) then if not fTOBPieceRG.InsertDB(nil) then V_PGI.IOError := OeUnknown;
  if (V_PGI.ioerror = OeOk) and (fTOBBasesRg.detail.count > 0) then if not fTOBBasesRg.InsertDB(nil) then V_PGI.IOError := OeUnknown;
end;


procedure TnewPiece.SetAffaire(Affaire: string);
var Aff0, Aff1, Aff2, Aff3, avenant: string;
begin
	Aff0 := copy(Affaire,1,1);
	BTPCodeAffaireDecoupe (affaire,aff0,Aff1,Aff2,Aff3,avenant,tacreat,false);
  fTOBPiece.putValue('GP_AFFAIRE',affaire);
  fTOBPiece.putValue('GP_AFFAIRE1',aff1);
  fTOBPiece.putValue('GP_AFFAIRE2',aff2);
  fTOBPiece.putValue('GP_AFFAIRE3',aff3);
  fTOBPiece.putValue('GP_AVENANT',avenant);
end;

procedure TnewPiece.SetDatePiece(DateCreat: TDateTime);
begin
	fTOBPiece.putValue('GP_DATEPIECE',DateCreat);
end;

procedure TnewPiece.ConstitueContrat(AffaireContrat,Affaire: string);
var Aff0,Aff1,Aff2,Aff3,Avenant,theAuxiliaire : string;
		DateFin : TdateTime;
    QQ : Tquery;
    NbMois,NbIt : integer;
    MtEcheance : double;
begin
	DateFin := Idate2099;
	Aff0 := Copy(AffaireContrat,1,1);
	BTPCodeAffaireDecoupe (AffaireContrat,aff0,Aff1,Aff2,Aff3,avenant,tacreat,false);
  fTOBAffaire.InitValeurs(false);
  fTOBAffaire.putValue ('AFF_AFFAIRE',AffaireContrat);
  fTOBAffaire.putValue ('AFF_AFFAIRE0',Aff0);
  fTOBAffaire.putValue ('AFF_AFFAIRE1',Aff1);
  fTOBAffaire.putValue ('AFF_AFFAIRE2',Aff2);
  fTOBAffaire.putValue ('AFF_AFFAIRE3',Aff3);
  fTOBAffaire.putValue ('AFF_AVENANT',Avenant);
  fTOBAffaire.putValue ('AFF_TIERS',ftobPiece.GetValue('GP_TIERS') );
  fTOBAffaire.putValue ('AFF_STATUTAFFAIRE','INT' );
  fTOBAffaire.putValue ('AFF_ETATFFAIRE','ENC' );
  fTOBAffaire.putValue ('AFF_CREATEUR',V_PGI.User );
  fTOBAffaire.putValue ('AFF_UTILISATEUR',V_PGI.User );
  fTOBAffaire.putValue ('AFF_DATECREATION',now );
  fTOBAffaire.putValue ('AFF_DATEMODIF',IDate2099 );
  fTOBAffaire.putValue ('AFF_LIBELLE','Contrat de maintenance progiciels');
  fTOBAffaire.putValue ('AFF_DATEREPONSE',IDate1900);
  fTOBAffaire.putValue ('AFF_DATESIGNE',fDateAcc);
  fTOBAffaire.putValue ('AFF_DATEDEBUT',fDateDebut);
  (*
  QQ:=OpenSQL('SELECT BRE_TYPEACTION,BRE_NBMOIS FROM BRECONDUCTION WHERE BRE_CODE="'+VH_GC.AFFReconduction+'"', TRUE);
  if not QQ.eof then
  begin
    if QQ.findField('BRE_TYPEACTION').asString = '' then
    begin
      nbMois := QQ.findField ('BRE_NBMOIS').asInteger;
      DateFin := IncMonth(fTOBAffaire.GetValue('AFF_DATEDEBUT'), NbMois);
      DateFin := PlusDate (DateFin,-1,'J');
    end;
  end;
  ferme (QQ);
  *)
  nbMois := GetParamSocSecur('SO_BTDUREEENGAGEMENT',12);
  DateFin := IncMonth(fTOBAffaire.GetValue('AFF_DATEDEBUT'), NbMois);
  DateFin := PlusDate (DateFin,-1,'J');
  fTOBAffaire.putValue ('AFF_DATEFIN', DateFin) ;
  NbMois:= StrToInt(GetParamSoc ('SO_AFCALCULFIN'));
  fTOBAffaire.putValue ('AFF_DATELIMITE', IncMonth(DateFin, NbMois)) ;
  NbMois:= StrToInt(GetParamSoc ('SO_AFALIMCLOTURE'));
  fTOBAffaire.putValue('AFF_DATECLOTTECH', IncMonth(DateFin, NbMois));
  NbMois:= StrToInt(GetParamSoc ('SO_AFCALCULGARANTI'));
  fTOBAffaire.putValue('AFF_DATEGARANTIE', IncMonth(DateFin, NbMois));
  fTOBAffaire.putValue ('AFF_DATERESIL', idate2099) ;
  fTOBAffaire.putValue ('AFF_DEVISE', ftobPiece.getValue('GP_DEVISE')) ;
  fTOBAffaire.putValue ('AFF_SAISIECONTRE','-') ;
  fTOBAffaire.putValue ('AFF_TOTALHT',ftobPiece.getValue('GP_TOTALHT'));
  fTOBAffaire.putValue ('AFF_TOTALHTDEV',ftobPiece.getValue('GP_TOTALHTDEV'));
  fTOBAffaire.putValue ('AFF_TOTALTTC',ftobPiece.getValue('GP_TOTALTTC'));
  fTOBAffaire.putValue ('AFF_TOTALTTCDEV',ftobPiece.getValue('GP_TOTALTTCDEV'));
  fTOBAffaire.putValue ('AFF_TOTALTAXE',ftobPiece.getValue('GP_TOTALTTC')-ftobPiece.getValue('GP_TOTALHT'));
  fTOBAffaire.putValue ('AFF_TOTALTAXEDEV',ftobPiece.getValue('GP_TOTALTTCDEV')-ftobPiece.getValue('GP_TOTALHTDEV'));
  fTOBAffaire.putValue ('AFF_TOTALHTGLO',ftobPiece.getValue('GP_TOTALHT'));
  fTOBAffaire.putValue ('AFF_TOTALHTGLODEV',ftobPiece.getValue('GP_TOTALHTDEV'));
  fTOBAffaire.putValue ('AFF_CALCTOTHTGLO','X');
  fTOBAffaire.putValue ('AFF_GENERAUTO','CON');
  fTOBAffaire.putValue ('AFF_PERIODICITE', GetParamSocSecur('SO_AFPERIODICTE', 'M')) ;
  fTOBAffaire.putValue ('AFF_MULTIECHE', '-') ;
  fTOBAffaire.putValue ('AFF_TERMEECHEANCE', GetParamSocSecur('SO_AFTERMEECHE', '')) ;
  fTOBAffaire.putValue ('AFF_METHECHEANCE', GetParamSocSecur('SO_AFMETHECHE', '')) ;
  fTOBAffaire.putValue ('AFF_DATEDEBGENER',fDateDebut) ;
  fTOBAffaire.putValue ('AFF_DATEFINGENER',DateFin) ;
  fTOBAffaire.putValue ('AFF_DETECHEANCE','DMM') ;  // montant doc = montant d'un mois
  fTOBAffaire.putValue ('AFF_INTERVALGENER',	GetParamSocSecur('SO_AFINTERVAL',1));
  fTOBAffaire.putValue ('AFF_REPRISEACTIV',	'TOU');
  fTOBAffaire.putValue ('AFF_ETATAFFAIRE',	'ACP');

  //Modif FV : Gestion des contrats au nombre d'interventions
  if fTOBAffaire.GetValue ('AFF_PERIODICITE') = 'NBI' then
    NbIt := 1
  else
    NbIt := EvaluationNbEcheances (fTOBAffaire.GetValue ('AFF_AFFAIRE'),
     															 fTOBAffaire.GetValue ('AFF_PERIODICITE'),
                                   fTOBAffaire.GetValue ('AFF_INTERVALGENER'), fDateDebut, DateFin) ;

  MtEcheance := CalculMtEcheanceContrat ( fTOBAffaire.GetValue ('AFF_TOTALHTDEV'),
  																				fTOBAffaire.GetValue ('AFF_DETECHEANCE'),
  																				fTOBAffaire.GetValue ('AFF_PERIODICITE'),
                                          fTOBAffaire.getValue ('AFF_INTERVALGENER'),
                                          nbIt);
  fTOBAffaire.putValue ('AFF_MONTANTECHEDEV',MtEcheance);
  fTOBAffaire.putValue ('AFF_MONTANTECHE',DeviseToEuro (MtEcheance, DEV.Taux, DEV.Quotite)) ;
  fTOBAffaire.putValue ('AFF_NATUREPIECEG', 'FPR') ;
  fTOBAffaire.putValue ('AFF_PROFILGENER', GetParamSocSecur('SO_AFPROFILGENER', '')) ;
  fTOBAffaire.putValue ('AFF_RECONDUCTION', VH_GC.AFFReconduction) ;
  fTOBAffaire.putValue ('AFF_COEFFREVALO', 1) ;
  fTOBAffaire.putValue ('AFF_ETABLISSEMENT', ftobPiece.getValue('GP_ETABLISSEMENT')) ;
  fTOBAffaire.putValue ('AFF_DOMAINE', ftobPiece.getValue('GP_DOMAINE')) ;
  fTOBAffaire.putValue ('AFF_CREERPAR', 'SAI') ;
  fTOBAffaire.putValue ('AFF_ADMINISTRATIF', '-') ;
  fTOBAffaire.putValue ('AFF_MODELE', '-') ;
  fTOBAffaire.putValue ('AFF_AFFAIREHT', 'X') ;
  fTOBAffaire.putValue ('AFF_TYPEPREVU', 'GLO') ;
  fTOBAffaire.putValue ('AFF_REGROUPEFACT', 'AUC') ;
  fTOBAffaire.putValue ('AFF_AFFAIREREF', fTOBAffaire.GetValue ('AFF_AFFAIRE') ) ;
  fTOBAffaire.putValue ('AFF_ISAFFAIREREF', '-') ;
  fTOBAffaire.putValue ('AFF_AFFCOMPLETE', 'X') ;
  fTOBAffaire.putValue ('AFF_MULTIPIECES', '-') ;
  // AFF_DESCRIPTIF
  fTOBAffaire.putValue ('AFF_CHANTIER', Affaire) ;
  if  fTOBPiece.getValue('GP_TIERSFACTURE') <> '' then
  begin
    QQ := OpenSql ('SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="'+fTOBPiece.getValue('GP_TIERSFACTURE')+'" AND T_NATUREAUXI IN ("CLI","PRO")',true,1,'',true);
    if not QQ.eof then theAuxiliaire := QQ.findField('T_AUXILIAIRE').asstring else theAuxiliaire := '';
    ferme (QQ);
  	fTOBAffaire.putValue ('AFF_FACTURE', theAuxiliaire) ;
  end else
  begin
    QQ := OpenSql ('SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="'+fTOBPiece.getValue('GP_TIERS')+'" AND T_NATUREAUXI IN ("CLI","PRO")',true,1,'',true);
    if not QQ.eof then theAuxiliaire := QQ.findField('T_AUXILIAIRE').asstring else theAuxiliaire := '';
    ferme (QQ);
  	fTOBAffaire.putValue ('AFF_FACTURE', theAuxiliaire) ;
  end;
  GenereEcheancesContrat;
end;

procedure TnewPiece.GenereEcheancesContrat;
var DateDebut,DateFin,DateDebutFac,DateLiquid,DateResil : Tdatetime;
		TypeCalcul,ModeMontantDoc,MethodeCalcul,TermeEche : string;
    MtEcheance,MtGlobal : double;
begin
	ModeMontantDoc := fTOBAffaire.GetValue('AFF_DETECHEANCE');
  MethodeCalcul := fTOBAffaire.GetValue('AFF_METHECHEANCE');
  DateDebutfac := fTOBAffaire.GetValue('AFF_DATEDEBGENER');
  TermeEche := fTOBAffaire.GetValue('AFF_TERMEECHEANCE');
  MtEcheance := fTOBAffaire.GetValue ('AFF_MONTANTECHEDEV');
  MtGlobal := fTOBAffaire.GetValue ('AFF_TOTALHTGLODEV');

  TypeCalcul := fTOBAffaire.GetValue  ('AFF_PERIODICITE') ;
  DateDebut := fTOBAffaire.GetValue ('AFF_DATEDEBGENER');
  DateLiquid :=fTOBAffaire.GetValue ('AFF_DATEFACTLIQUID')  ;
  DateResil := fTOBAffaire.GetValue ('AFF_DATERESIL')  ;

  if fTOBAffaire.GetValue ('AFF_METHECHEANCE') = 'CIV' then
  begin
  	DateDebut := GetDateDebutPeriode (TypeCalcul, DateDebut,
    															 fTOBAffaire.GetValue ('AFF_INTERVALGENER') ,
                                   fTOBAffaire.GetValue ('AFF_METHECHEANCE') ) ;
  end;

  DateFin := fTOBAffaire.GetValue ('AFF_DATEFINGENER') ;

	ConstitueTOBEcheances (fTOBFACTAFF,MtGlobal,
                         fTOBAffaire.GetValue ('AFF_PROFILGENER'),
                         fTOBAffaire.GetValue ('AFF_AFFAIRE') ,
                         fTOBAffaire.GetValue ('AFF_PERIODICITE'),
                         fTOBAffaire.GetValue ('AFF_REPRISEACTIV'),
                         fTOBAffaire.GetValue ('AFF_TIERS'),
                         fTOBAffaire.GetValue  ('AFF_GENERAUTO'),
                         [tmaMntEch],
                         fTOBAffaire.GetValue  ('AFF_INTERVALGENER'),
                         MtEcheance,
                         DateDebut,DateFin,DateLiquid,Idate1900,Idate2099,DateResil,
                         DEV,False,true,false,fTOBAffaire.GetValue('AFF_MULTIECHE'),
												 ModeMontantDoc,MethodeCalcul,DateDebutFac,termeEche);
end;


procedure TnewPiece.ConstitueCodeContrat(var AffaireContrat: string);
var Aff0,Aff1,Aff2,Aff3,Avenant : string;
		QQ : TQuery;
begin
	Aff0 := 'I';
  Aff1 := '';
  Aff3 := '';
  Avenant := '';
  QQ := OpenSql('SELECT ET_CHARLIBRE1 FROM ETABLISS WHERE ET_ETABLISSEMENT="'+
  							fTOBPiece.getValue ('GP_ETABLISSEMENT')+'"',true,1,'',true);
	if not QQ.eof then
  begin
  	Aff2 := QQ.findField('ET_CHARLIBRE1').AsString;
  end;
  ferme (QQ);
  if Aff2 = '' then
  begin
  	raise Exception.Create('L''etablissement n''est pas rensigné dans le document');
  end;
	AffaireContrat := CodeAffaireRegroupe (Aff0,Aff1,Aff2,Aff3,Avenant,tacreat,true,false,true);
end;

procedure TnewPiece.DefiniAffairePiece;
var QQ : TQuery;
    CodessAff,Part0,Part1,Part2,Part3,Avenant : string;
    Numero : integer;
begin
  QQ := OpenSql ('SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE='+
  			'(SELECT GP_AFFAIREDEVIS FROM PIECE WHERE GP_NATUREPIECEG="'+
        fOldnaturePieceg+'" AND GP_SOUCHE="'+
        fOldSouche+'" AND GP_NUMERO='+
        IntToStr(fOldNumero)+' AND GP_INDICEG='+
        IntToStr(FoldIndiceg)+')' ,true,1,'',true);
  if not QQ.eof then
  begin
  	fTOBAffaireP.selectdb('',QQ);
    Numero := fTOBPiece.getValue('GP_NUMERO');
    Codessaff := Format ('%.8d',[Numero]);
    Codessaff := 'Z'+fTOBPiece.getValue('GP_NATUREPIECEG')+fTOBPiece.getValue('GP_SOUCHE')+Codessaff+'00';
    fTOBAffaireP.putValue('AFF_AFFAIRE',CodessAff);
    //
    BTPCodeAffaireDecoupe(Codessaff, Part0, Part1, Part2, Part3,Avenant,taConsult,False);
    fTOBAffaireP.PutValue('AFF_AFFAIRE', Codessaff);
    fTOBAffaireP.PutValue('AFF_AFFAIRE0', 'Z');
    fTOBAffaireP.PutValue('AFF_AFFAIRE1', Part1);
    fTOBAffaireP.PutValue('AFF_AFFAIRE2', Part2);
    fTOBAffaireP.PutValue('AFF_AFFAIRE3', Part3);
    fTOBAffaireP.PutValue('AFF_AVENANT', Avenant);
    fTOBAffaireP.PutValue('AFF_AFFAIREINIT', fTOBPiece.getValue('GP_AFFAIRE'));
    fTOBAffaireP.PutValue('AFF_DATECREATION', fTOBPiece.getValue('GP_DATEPIECE'));
    fTOBpiece.PutValue('GP_AFFAIREDEVIS',CodessAff);
  end;
  ferme (QQ);
end;

{ TlisTnewPiece }

function TlisTnewPiece.Add(AObject: TnewPiece): Integer;
begin
  result := Inherited Add(AObject);
end;

procedure TlisTnewPiece.clear;
var Indice : integer;
begin
  for Indice := 0 to Count -1 do
  begin
    TnewPiece(Items [Indice]).free;
  end;
  inherited;
end;

destructor TlisTnewPiece.destroy;
begin
  clear;
  inherited;
end;

function TlisTnewPiece.findPiece(TypePiece: TTypePiece): TnewPiece;
var Indice : integer;
begin
  result := nil;
  for Indice := 0 to Count -1 do
  begin
    if Items[Indice].TypePiece = TypePiece then
    begin
      result:=Items[Indice];
      break;
    end;
  end;
end;

function TlisTnewPiece.GetItems(Indice: integer): TnewPiece;
begin
  result := TnewPiece (Inherited Items[Indice]);
end;

procedure TlisTnewPiece.SetItems(Indice: integer; const Value: TnewPiece);
begin
  Inherited Items[Indice]:= Value;
end;

procedure LSEAcceptEtude( parms: array of variant; nb: integer ) ;
const Messagetext = 'Confirmez-vous la validation des études sélectionnées ?';
var F : TForm;
		Q : TQuery ;
		i :integer;
		TOBAff,TOBPieces,TOBPiece :TOB;
		Numero, Indice : integer;
		NaturePiece, Souche :string;
		Mess,Sql : string;
begin
  F:=TForm(Longint(Parms[0]));
  if (TFMul(F).FListe=nil) then exit;
  Q:=TFMul(F).Q;
  Mess:=Messagetext;
  if (PGIAsk (Mess, F.Caption)<>mrYes) then exit;
  SourisSablier;

  TOBAff:=TOB.Create('AFFAIRE',Nil,-1) ;

  // on crée une TOB de toutes les études sélectionnées
  TobPieces := TOB.Create ('Liste des études',NIL, -1);

  try
    if TFMul(F).Fliste.AllSelected then
    BEGIN
      Q.First;
      while Not Q.EOF do
      BEGIN
        NaturePiece:=Q.FindField('GP_NATUREPIECEG').AsString;
        Souche:=Q.FindField('GP_SOUCHE').AsString;
        Numero:=Q.FindField('GP_NUMERO').AsInteger;
        Indice:=Q.FindField('GP_INDICEG').AsInteger;
        // Remplissage de la TOB contenant les pieces selectionnées
        TobPiece := Tob.Create ('ONE PIECE', TobPieces,-1);
        Sql := NaturePiece
               + ';'+ Souche
               + ';'+ inttostr(Numero)
               + ';'+ inttostr(Indice);
        TOBPiece.AddChampSupValeur('CLEDOC',Sql);
        Q.NEXT;
      END;
      TFMul(F).Fliste.AllSelected:=False;
    END else
    begin
      for i:=0 to TFMul(F).Fliste.nbSelected-1 do
      begin
        TFMul(F).Fliste.GotoLeBookmark(i);
        NaturePiece:=TFMul(F).Fliste.datasource.dataset.FindField('GP_NATUREPIECEG').AsString;
        Souche:=TFMul(F).Fliste.datasource.dataset.FindField('GP_SOUCHE').AsString;
        Numero:=TFMul(F).Fliste.datasource.dataset.FindField('GP_NUMERO').AsInteger;
        Indice:=TFMul(F).Fliste.datasource.dataset.FindField('GP_INDICEG').AsInteger;
        TobPiece := Tob.Create ('ONE PIECE', TobPieces,-1);
        Sql := NaturePiece
               + ';'+ Souche
               + ';'+ inttostr(Numero)
               + ';'+ inttostr(Indice);
        TOBPiece.AddChampSupValeur('CLEDOC',Sql);
    	end;
    end;
    LSEGenerePiecesGestion (TobPieces);
  finally
    TOBAff.Free; TobPieces.Free;
    SourisNormale ;
	end;
end;

{ TlistDesPieces }

function TlistDesPieces.Add(AObject: TnewPieceGestion): Integer;
begin
  result := Inherited Add(AObject);
end;

procedure TlistDesPieces.clear;
var indice : integer;
begin
  inherited;
  for Indice := 0 to Count -1 do
  begin
    TnewPieceGestion(Items [Indice]).free;
  end;
end;

destructor TlistDesPieces.destroy;
begin
	clear;
  inherited;
end;

function TlistDesPieces.GetItems(Indice: integer): TnewPieceGestion;
begin
  result := TnewPieceGestion (Inherited Items[Indice]);
end;

procedure TlistDesPieces.SetItems(Indice: integer;
  const Value: TnewPieceGestion);
begin
  Inherited Items[Indice]:= Value;
end;

Initialization
RegisterAglProc( 'LSEAcceptEtude', TRUE , 0, LSEAcceptEtude);
end.
