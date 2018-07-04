{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 24/09/2004
Modifié le ... : 24/09/2004
Description .. : - BPY le 24/09/2004 - Passage en eAGL
Mots clefs ... :
*****************************************************************}

unit VerCpta;

//Document Intégrateur n° 60017_86.doc (G:\client\0\7)
// Info à garder !! ...
//      Control sur les pièces en Devise vis-à-vis de la Quotité !
//
// Test Sur TotDebit, Totcredit des Cptes Genes ??
//
// Test sur Equilibre bloquant ?
// Test sur TAUXDEV en ECC suppr ?
// Test sur TAUXDEV en ECC suppr ?

//================================================================================
// Interface
//================================================================================
interface

uses
    Windows,
    Messages,
    SysUtils,
    Classes,
    Graphics,
    Controls,
    Forms,
    Dialogs,
    Mask,
    StdCtrls,
    Hctrls,
    Hcompte,
    Buttons,
    ExtCtrls,
    Ent1,
    HEnt1,
    UTILEDT,
{$IFDEF EAGLCLIENT}
{$ELSE}
    DB,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
    HQry,
{$ENDIF}
{$IFDEF VER150}
   Variants,
 {$ENDIF}
    UTOB,
    UlibExercice,
    hmsgbox,
    SaisUtil,
    HStatus,
    RapSuppr,
    RappType,
    ComCtrls,
    CRITEDT,
    HSysMenu,
    ParamDat,
    CpteSav,
    UtilSais,
    {$IFDEF MODENT1}
    CPTypeCons
    {$ELSE}
    tCalcCum
    {$ENDIF MODENT1}
    ;

//==================================================
// Definition de types
//==================================================
type
    TTypeVerif = (tvEcr, tvEcrAna, tvEcrbud);

type
    TEnregInfos = class
        Inf1 : string;
        Inf2 : string;
        Inf3 : string;
        Inf4 : string;
        Inf5 : string;
        Inf6 : string; // FQ 12405
        Inf7 : string; // FQ 12405
        D1 : TDateTime;
        D2 : TDateTime;
        Nb : Byte;
        Bol1, Bol2, Bol3 : Boolean;
        SuperBol : T5B;
        Mont : Double;
    end;

type
    TINFOSCPTE = record
        General : string;
        NatureGene : string;
        Collectif, Lettrable, Pointable : Boolean;
        Ventilable : T5B;
    end;

type
    TINFOSAUXI = record
        Auxi : string;
        NatureAux : string;
        Lettrable : Boolean;
    end;

type
    TErrLetDev = class
        Lettres : string;
        Auxi : string;
        Gene : string;
        EtatLet : string;
    end;

type
    TINFOSMVTANA = record
        General : string;
        AnalPur : Boolean;
        TotEcr, TotDev : Double;
        TypeANO, EcrANO : string;
    end;

type
    TINFOSAUTRES = record
        J_Axe : string;
        J_COMPTEINTERDIT : string;
        J_NATUREJAL : string;
        J_MODESAISIE : string;
        J_VALIDEEN : string; // FQ 12405
        J_VALIDEEN1 : string; // FQ 12405
        EX_DATEEXO : array[0..1] of TDateTime;
    end;

type
    TINFOSPIECE = record
        Jal, Exo : string;
        Dat : TDateTime;
        NumP : Integer;
    end;

type
    TERRORIMP = class
        NumChrono : Integer;
        NumErreur : string;
    end;

type
    tSetAxe = set of Byte;

//==================================================
// Externe
//==================================================

procedure VerifCpta;
function VerPourClo(Exer : TExoDate) : Boolean; { Verif avant cloture }
function VerPourImp(var ListeError : HTStringList;EnBatch : boolean;TypeVerif : TTypeVerif) : Boolean; { Vérif pour Import }
function VerPourImp2(ListeEntetePieceFausse : HTStringList;ListePbPiece : TList;EnBatch : boolean;TypeVerif : TTypeVerif) : Boolean; { Vérif pour Import }
function VerPourImp3(ListeEntetePieceFausse, ListePbAna : HTStringList;ListePbPiece : TList;EnBatch : boolean;TypeVerif : TTypeVerif;ShuntPbAna : Boolean;var PbAna : Boolean) : Boolean; { Vérif pour Import }
function VerPourImpRecupSISCOPGI(ListeEntetePieceFausse, ListePbAna : HTStringList;ListePbPiece : TList;EnBatch : boolean;TypeVerif : TTypeVerif;ShuntPbAna : Boolean;var PbAna : Boolean) : Boolean; { Vérif pour Import }
procedure ControleMvt(Quel : Byte);

procedure VerifAno(Co : Boolean); {Verif des ANouveau }
function LeSolde(D, C : Double) : Double;
function VirguleToPoint(Montant : Double) : string;

//==================================================
// Definition de class
//==================================================
type
    TFVerCpta = class(TForm)
        HPB : TPanel;
    BValider: THBitBtn;
    BFerme: THBitBtn;
        Panel1 : TPanel;
        TFJal : THLabel;
        FJal1 : THCpteEdit;
        TFaJ : TLabel;
        FJal2 : THCpteEdit;
        FJalJoker : TEdit;
        TFJalJoker : THLabel;
        TFExercice : THLabel;
        FExercice : THValComboBox;
        FDateCpta1 : TMaskEdit;
        TFDateCpta1 : THLabel;
        TFDateCpta2 : TLabel;
        FDateCpta2 : TMaskEdit;
        TFEtab : THLabel;
        FEtab : THValComboBox;
        MsgRien : THMsgBox;
        MsgLibel : THMsgBox;
        MsgLibel2 : THMsgBox;
        TFTypeEcriture : THLabel;
        FTypeEcriture : THValComboBox;
        Shape1 : TShape;
        MsgBar : THMsgBox;
        TTravail : TLabel;
        Panel2 : TPanel;
        TNBError1 : TLabel;
        TNBError2 : TLabel;
        TNBError3 : TLabel;
    BStop: THBitBtn;
        TFVerification : THLabel;
        FVerification : TComboBox;
        TNBError5 : TLabel;
        TNBError4 : TLabel;
        TCEtab : THValComboBox;
        TCModP : THValComboBox;
        TCNatPiece : TComboBox;
        TTemp : TLabel;
        FSolde : THNumEdit;
        HMTrad : THSystemMenu;
        TFNumPiece1 : THLabel;
        FNumPiece1 : TMaskEdit;
        TFNumPiece2 : TLabel;
        FNumPiece2 : TMaskEdit;
        CtrlDetGen : TCheckBox;
        CtrlEcrGen : TCheckBox;
        CtrlEcrAna : TCheckBox;
        CtrlDetAna : TCheckBox;
        TNbError6 : TLabel;
        CtrlAnaOff : TCheckBox;

        procedure FormCreate(Sender : TObject);
        procedure PrepareShow;
        procedure FormShow(Sender : TObject);
        procedure FormClose(Sender : TObject;var Action : TCloseAction);

        procedure BValiderClick(Sender : TObject);
        procedure BStopClick(Sender : TObject);

        procedure FJal1Change(Sender : TObject);
        procedure FExerciceChange(Sender : TObject);
        procedure FVerificationChange(Sender : TObject);

        procedure FDateCpta1KeyPress(Sender : TObject;var Key : Char);

    private
    { Déclarations privées }
        OkVerif, StopVerif, MAJ, DejaLance, LanceDirect : Boolean;
        LaListe : TList;
        Mvt : TInfoMvt;
        MvtBud : TInfoMvtBud;
        InfAutres : TINFOSAUTRES;
        NbEnreg, NbError : Integer;
        PourImport : Boolean;
        EnBatch : Boolean;
        CGen : TINFOSCPTE;
        CTiers : TINFOSAUXI;
        ErrAnal, ErrEcr, ErrMvt, ErrLet, ErrMvtA, ErrMvtB : Integer;
        FE, FY, FBE, FICECR, FICANA, FICBUDECR : string;
        VCrit, VCritLance : TCritEdt;
        LetListe : TStringList;
        ImpListe : HTStringList;
        ListeDEV : TStringList;
        ListeJAL : TStringList;
        ListeEXO : TStringList;
        ListeSec : TStringList;
        ListeJalBud : TStringList;
        ListeCptBud : TStringList;
        ListeSecBud : TStringList;
        ListeGen : TStringList;
        ListeBQE : TStringList;
        ListeAUX : TStringList;
        INFOSPIECE : TINFOSPIECE;
        ListePbAna : HtStringList;
        ShuntPbAna, PbAna : Boolean;

        QSum : TQuery;
        QEcr : TQuery;
        QAnal : TQuery;
        QInfoAnaG : TQuery;
        QLettrage : TQuery;
        QEquilLet : TQuery;
        QLetDev : TQuery;
        QAnaGene : TQuery;
        QSumGeAn : TQuery;
        QEcrBud : TQuery;

        // pour savoir s'il y des enreg a traité et commance le chargement !
        function SiEnreg : Boolean;
        function CompteLettrage : Boolean;
        // Charge et libere des info complemantaire au traitement
        procedure ChargeInfos;
        procedure VideInfos;
        // initialisation des libelle !
        procedure InitLabel;
        // verif de validité des date
        function DateOk : Boolean;
        // liberation de la liste de lettrage ??? et de chaqu'un de ces composant
        procedure VideLetListe;
        // recup des critere de test
        procedure RecupCrit;

        // Traitement de verification principal !
        procedure LanceVerif;
        procedure lanceVerifAnaOff;
        procedure Reparation;

        // interuption du traitement
        function TestBreak : Boolean;

        // correction ??
        procedure Corrige(Q : TQuery;Champ, Valeur : string);

        // recup d'information
        procedure InfosCptGen(Ge : string); // compte generaux
        procedure InfosCptAux(Au : string); // compte auxiliaire

// traitement des ecriture
        // etape du traitement
        procedure SQLEcr;
        procedure SqlEquilEcr(Ec : Byte);
        // svg de l'enreg d'ecriture dans une structure mvt
        procedure EnregMvtEcr;
        // equilibrage de la piece
        procedure EnregEquil(Q : TQuery;Quelle : byte);
        function EquilEcr : Boolean;
        procedure ErrEquilEcr(var Err : Boolean;E : Byte;D, C, DD, CD, DE, CE : Double);

// traitement de l'analytique
        // etape du traitement
        procedure SqlAnal;
        // svg de l'enreg analytique dans une structure mvt
        procedure EnregMvtAna;
        // equilibrage de la piece analytique
        function EquilAnal : Boolean;
        procedure VerifAutres(var IFA : TINFOSMVTANA);
        procedure EnregEquilA;
        // correction de l'equilibre analytique
        procedure CorrEquilAna(E : Byte;M : TabTot);

// traitement du lettrage
        // etape du traitement
        procedure SQLLet;
        procedure SQLVerifLet(Pourqui : Byte);
        function SQLEquilLet : Boolean;
        // svg de l'enreg analytique dans une structure mvt
        procedure EnregLet(Q : TQuery);

// traitement du budget
        // etape du traitement
        procedure SQLEcrBud;
        // svg de l'enreg analytique dans une structure mvt
        procedure EnregMvtEcrBud;

        // ajout des info de la pieces dans une liste
        procedure lalisteAdd(G : DelInfo);

        // ???
        function GoListe1(Entite : TObject;Quel, Rem : string;I : Byte;CodeIMP : string) : DelInfo;
        function GoListe2(Entitee : TInfoMvt;Quel, Rem : string;F : TField;CodeIMP : string) : DelInfo;

        // capitalization ???
        function Majuscule(St : string) : Boolean;

        // gestion des libelle d'affichage du travail et des libelle d'erreur !
        procedure TuFaisQuoi(Q : TQuery;DeQuoi : Byte);

        // test d'existance de la valeur dans la base
        function Existe(V : Byte;St : string) : Boolean; // devise, journal ou nature de pieces
        function OkExo(E : String3;D : TDateTime) : Boolean; // exercice
        procedure TrouveDecDev(St : String3;var Dec : Byte); // devise
        function ExisteSec(Ax, Cod : string) : Boolean; // section
        function OkPourSum : Boolean; // lettrage
        function ExisteCpteBud(cod : string) : boolean; // compte budgetaire
        function ExisteJalBud(cod : string) : boolean; // journal budgetaire
        function ExisteSecBud(Ax, Cod : string) : Boolean; // section budgetaire

        // verification de coherence
        function VerifCompte(var let, Poin, vent : Boolean) : Boolean; // compte gene et auxi
        function VerifInfoJal : Boolean; // journal
        function VerifInfoPerio : Boolean; // periode
        function VerifInfoExo : Boolean; // exercice
        function VerifInfoEcrAno : Boolean; // ecriture d'a nouveau
        function VerifInfoEche(Lettrable, pointable : Boolean) : Boolean; // echeance
        function VerifInfoDevise : Boolean; // devise
        function VerifInfoPieces : Boolean; // piece
        function VerifTauxDev(MvtEcr : Boolean) : Boolean; // taux de la devise
        function VerifEtab : Boolean; // etablissement
        function EcrAnoOk(Journal, Ecran : string) : Boolean; // journal des ecriture d'a nouveau

        // verification de coherence analytique
        function VerifInfoAnal(Ventilable : Boolean) : Boolean; // existance de l'analytique
        function VerifCptAnal : Boolean; // Section axe et compte gene
        function VerifInfoJalAnal : Boolean; // journal
        function VerifInfoExoAnal : Boolean; // exercice
        function VerifInfoEcrAnoAnal : Boolean; // ecriture d'a nouveau
        function VerifInfoPieceAnal : Boolean; // piece
        function VerifInfoDeviseAnal : Boolean; // devise
        function VerifEtabAnal : Boolean; // etablissement

        // verification de coherence du lettrage
        function VerifMvtNoLet : Boolean; // etat du lettrage
        function VerifMvtLet : Boolean; // couverture selon le lettrage
        function VerifLetDev : Boolean; // devise

        // verification de coherence budgetaire
        function VerifCompteBud : Boolean; // compte budgetaire
        function VerifInfoJalBud : Boolean; // journal
        function VerifInfoExoBud : Boolean; // exercice
        function VerifInfoSecBud : Boolean; // section
        function VerifEtabBud : Boolean; // etablissement

//**************** non utilisé
//        Function  VerifInfoConfAnal : Boolean ;
//        Function  VerifInfoConf : Boolean ;
//        procedure GoListeImp(Chrono : Integer ; Code : String) ;
//        procedure MAJMVT(Champ, Valeur : String) ;
//        function  DevBqe(Cpt : String ) : String ;
//        function  DonneQuotite(Dev : String) : Double ;
    public
    { Déclarations publiques }
    end;

//================================================================================
// Implementation
//================================================================================
implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  TImpFic
    ;

//==================================================
// Definition des Constante
//==================================================
const
    MaxNbErreur = 999999;

//==================================================
// Definition des variables
//==================================================
var
    PourNewImport : Boolean;

{$R *.DFM}

//==================================================
// fonctions d'appele
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure VerifCpta;
var
    VCpta : TFVerCpta;
begin
    PourNewImport := FALSE;
    VCpta := TFVerCpta.Create(Application);
    try
        VCpta.FVerification.ItemIndex := 0;
        VCpta.EnBatch := False;
        VCpta.ShuntPbAna := False;
        VCpta.PbAna := False;
        VCpta.ShowModal;
    finally
        VCpta.Free;
    end;
    Screen.Cursor := SyncrDefault;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function VerPourClo(Exer : TExoDate) : Boolean;
var
    VCpta : TFVerCpta;
begin
    PourNewImport := FALSE;
    VCpta := TFVerCpta.Create(Application);
    VCpta.EnBatch := False;
    VCpta.ShuntPbAna := False;
    VCpta.PbAna := False;
    try
        with VCpta do
        begin
            OkVerif := True;
            PourImport := False;
            FVerification.ItemIndex := 1;
            MAJ := False;
            FE := 'E';
            FY := 'Y';
            FICECR := 'ECRITURE';
            FICANA := 'ANALYTIQ';
            Fillchar(VCrit, SizeOf(VCrit), #0);
            with VCrit do
            begin
                Etab := '';
                Joker := False;
                Cpt1 := '';
                Cpt2 := '';
                Exo.Code := Exer.Code;
                Date1 := Exer.Deb;
                Date2 := Exer.Fin;
                QualifPiece := 'N';
                GL.NumPiece1 := 0;
                GL.NumPiece2 := 999999999;
            end;
            StopVerif := False;
            DejaLance := False;
            LanceDirect := False;
            if SiEnreg then
            begin
                ChargeInfos;
                lanceVerif;
                VideInfos;
            end;
            if not OkVerif then RapportdErreurMvt(Laliste, 3, MAJ, False);
            Result := OkVerif;
        end;
    finally
        VCpta.Free;
    end;
    Screen.Cursor := SyncrDefault;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function VerPourImp(var ListeError : HTStringList;EnBatch : boolean;TypeVerif : TTypeVerif) : Boolean;
var
    VCpta : TFVerCpta;
begin
    ListeError.Clear;
    PourNewImport := FALSE;
    VCpta := TFVerCpta.Create(Application);
    VCpta.EnBatch := EnBatch;
    VCpta.ShuntPbAna := False;
    VCpta.PbAna := False;
    try
        with VCpta do
        begin
            OkVerif := True;
            PourImport := True;
            case TypeVerif of
                tvEcr : FVerification.ItemIndex := 1;
                tvEcrAna :
                    begin
                        FVerification.ItemIndex := 1;
                        CtrlEcrGen.Checked := False;
                        CtrlDetGen.Checked := True;
                        CtrlEcrAna.Checked := False;
                        CtrlDetAna.Checked := True;
                    end;
                tvEcrBud :
                    begin
                        FVerification.ItemIndex := 3;
                        CtrlEcrGen.Checked := False;
                        CtrlDetGen.Checked := False;
                        CtrlEcrAna.Checked := False;
                        CtrlDetAna.Checked := False;
                    end;
            end;
            MAJ := False;
            FE := 'IE';
            FY := FE;
            FBE := FE;
            FICECR := 'IMPECR';
            FICANA := FICECR;
            FICBUDECR := FICECR;
            Fillchar(VCrit, SizeOf(VCrit), #0);
            with VCrit do
            begin
                Etab := '';
                Joker := False;
                Cpt1 := '';
                Cpt2 := '';
                QualifPiece := '';
                Date1 := IDate1900;
                Date2 := IDate2099;
                GL.NumPiece1 := 0;
                GL.NumPiece2 := 999999999;
            end;
            StopVerif := False;
            DejaLance := False;
            LanceDirect := False;
            if SiEnreg then
            begin
                ImpListe := HTStringList.Create;
                ImpListe.Clear;
                ChargeInfos;
                lanceVerif;
                VideInfos;
            end;
            if not OkVerif then RapportdErreurMvt(Laliste, 3, MAJ, EnBatch);
            ListeError := ImpListe;
            Result := OkVerif;
        end;
    finally
        VCpta.Free;
    end;
    Screen.Cursor := SyncrDefault;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function VerifPourImport(ListeEntetePieceFausse, ListePbAnal : HTStringList;ListePbPiece : TList;EnBatch : boolean;TypeVerif : TTypeVerif;ShuntPbAna : Boolean;var PbAna : Boolean) : Boolean; { Vérif pour Import }
var
    VCpta : TFVerCpta;
begin
    ListeEntetePieceFausse.Clear;
    PourNewImport := TRUE;
    VCpta := TFVerCpta.Create(Application);
    VCpta.EnBatch := EnBatch;
    VCpta.Laliste := ListePbPiece;
    VCpta.ShuntPbAna := ShuntPbAna;
    VCpta.PbAna := False;
    try
        with VCpta do
        begin
            OkVerif := True;
            PourImport := True;
            case TypeVerif of
                tvEcr : FVerification.ItemIndex := 1;
                tvEcrAna :
                    begin
                        FVerification.ItemIndex := 1;
                        CtrlEcrGen.Checked := False;
                        CtrlDetGen.Checked := True;
                        CtrlEcrAna.Checked := False;
                        CtrlDetAna.Checked := True;
                    end;
                tvEcrBud :
                    begin
                        FVerification.ItemIndex := 3;
                        CtrlEcrGen.Checked := False;
                        CtrlDetGen.Checked := False;
                        CtrlEcrAna.Checked := False;
                        CtrlDetAna.Checked := False;
                    end;
            end;
            if VH^.RecupSISCOPGI then
            begin
                CtrlEcrGen.Checked := False;
                CtrlDetGen.Checked := False;
                CtrlEcrAna.Checked := True;
                CtrlDetAna.Checked := False;
            end;
            MAJ := False;
            FE := 'IE';
            FY := FE;
            FBE := FE;
            FICECR := 'IMPECR';
            FICANA := FICECR;
            FICBUDECR := FICECR;
            Fillchar(VCrit, SizeOf(VCrit), #0);
            with VCrit do
            begin
                Etab := '';
                Joker := False;
                Cpt1 := '';
                Cpt2 := '';
                QualifPiece := '';
                Date1 := IDate1900;
                Date2 := IDate2099;
                GL.NumPiece1 := 0;
                GL.NumPiece2 := 999999999;
            end;
            StopVerif := False;
            DejaLance := False;
            LanceDirect := False;
            if SiEnreg then
            begin
                ImpListe := ListeEntetePieceFausse;
                ListePbAna := ListePbAnal;
                ChargeInfos;
                lanceVerif;
                VideInfos;
            end;
            if not OkVerif then RapportdErreurMvt(Laliste, 3, MAJ, EnBatch);
            Result := OkVerif;
        end;
    finally
        PbAna := VCpta.PbAna;
        VCpta.Free;
    end;
    Screen.Cursor := SyncrDefault;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function VerPourImp2(ListeEntetePieceFausse : HTStringList;ListePbPiece : TList;EnBatch : boolean;TypeVerif : TTypeVerif) : Boolean; { Vérif pour Import }
var
    pbAna : Boolean;
begin
    pbAna := FALSE;
    Result := VerifPourImport(ListeEntetePieceFausse, nil, ListepbPiece, EnBatch, TypeVerif, False, pbAna);
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function VerPourImp3(ListeEntetePieceFausse, ListePbAna : HTStringList;ListePbPiece : TList;EnBatch : boolean;TypeVerif : TTypeVerif;ShuntPbAna : Boolean;var PbAna : Boolean) : Boolean; { Vérif pour Import }
begin
    pbAna := FALSE;
    Result := VerifPourImport(ListeEntetePieceFausse, ListePbAna, ListepbPiece, EnBatch, TypeVerif, ShuntPbAna, pbAna);
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function VerPourImpRecupSISCOPGI(ListeEntetePieceFausse, ListePbAna : HTStringList;ListePbPiece : TList;EnBatch : boolean;TypeVerif : TTypeVerif;ShuntPbAna : Boolean;var PbAna : Boolean) : Boolean; { Vérif pour Recup SISCO PGI }
begin
    pbAna := FALSE;
    Result := VerifPourImport(ListeEntetePieceFausse, ListePbAna, ListepbPiece, EnBatch, TypeVerif, ShuntPbAna, pbAna);
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure ControleMvt(Quel : Byte);
var
    VCpta : TFVerCpta;
begin
    VCpta := TFVerCpta.Create(Application);
    try
        VCpta.EnBatch := False;
        VCpta.FVerification.ItemIndex := Quel;

        if VH^.EnSerie then
        begin
            VCpta.PrepareShow;
            VCPTA.BValider.Click;
        end
        else VCpta.ShowModal;

    finally
        VCpta.Free;
    end;
    Screen.Cursor := SyncrDefault;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure VerifAno(Co : Boolean); // CO == Correction prévoir de brancher la BAR
var
    Q, QE{, QLong} : TQuery;
    Err, ResultOk, Benefice : Boolean;
    ListErr : TList;
    X : DelInfo;
    CptR, LibR : string;
    Cha, Pro, Sol1, Sol2 : Double;
    Sum : TABTOT;
begin
    Showmessage(TraduireMemoire('Pas disponible pour l''instant'));
    Exit;

    { Attention --- V^  des compte de resultat pas fait dans ENT1 }

    ListErr := TList.Create;
    Err := False;
    { --------- Calcul du resultat N-1 pour report sur compte 12 en Benef ou Perte }
    ResultOk := True;
    Q := OpenSql(' select sum (e_debit-e_credit) as Cha from ecriture '
        + ' Left Join GENERAUX on E_GENERAL=G_GENERAL '
        + ' where e_exercice="' + VH^.Precedent.Code + '" and e_qualifpiece="N" and g_naturegene="CHA" ', True);
    if not Q.Eof then Cha := Arrondi(Q.FindField('Cha').AsFloat, V_PGI.OkdecV)
    else Cha := 0;
    Ferme(Q);
    Q := OpenSql(' select sum (e_debit-e_credit) as Pro from ecriture '
        + ' Left Join GENERAUX on E_GENERAL=G_GENERAL '
        + ' where e_exercice="' + VH^.Precedent.Code + '" and e_qualifpiece="N" and g_naturegene="PRO" ', True);
    if not Q.Eof then Pro := Arrondi(Q.FindField('Pro').AsFloat, V_PGI.OkdecV)
    else Pro := 0;
    Ferme(Q);
    Sol1 := LeSolde(Cha, Pro);
    Benefice := (Sol1 < 0);
    if Benefice then
    begin (*un truc comme V_PGI.OUVREBEN *)
        Q := OpenSql(' select G_GENERAL, G_LIBELLE, G_TOTDEBANO, G_TOTCREANO from generaux where g_general="12000000" ', True);
        CptR := Q.FindField('G_GENERAL').AsString;
        LibR := Q.FindField('G_LIBELLE').AsString;
        if Arrondi(Q.FindField('G_TOTCREANO').AsFloat, V_PGI.OkdecV) <> abs(Sol1) then ResultOk := False;
    end
    else
    begin (* un truc comme V_PGI.OUVREPER *)
        Q := OpenSql(' select G_GENERAL, G_LIBELLE, G_TOTDEBANO, G_TOTCREANO from generaux where g_general="12900000" ', True);
        CptR := Q.FindField('G_GENERAL').AsString;
        LibR := Q.FindField('G_LIBELLE').AsString;
        if Arrondi(Q.FindField('G_TOTDEBANO').AsFloat, V_PGI.OkdecV) <> abs(Sol1) then ResultOk := False;
    end;
    if not ResultOk then
    begin
        if CO then { Correction du report }
        begin
            if Benefice then
            begin
                ExecuteSql(' Update GENERAUX set G_TOTCREANO=' + STRFPOINT(abs(Sol1)) + ', G_TOTDEBANO=0 Where g_general="12000000" ');
                ExecuteSql(' Update GENERAUX set G_TOTDEBANO=0, G_TOTCREANO=0 Where g_general="12900000" ');
            end
            else
            begin
                ExecuteSql(' Update GENERAUX set G_TOTDEBANO=' + STRFPOINT(abs(Sol1)) + ', G_TOTCREANO=0 Where g_general="12900000" ');
                ExecuteSql(' Update GENERAUX set G_TOTDEBANO=0, G_TOTCREANO=0 Where g_general="12000000" ');
            end;
        end
        else
        begin
            X := DelInfo.Create;
            X.LeCod := CptR;
            X.LeLib := LibR;
            X.LeMess := FloatToStr(Q.FindField('G_TOTDEBANO').AsFloat);
            X.LeMess2 := FloatToStr(Q.FindField('G_TOTCREANO').AsFloat);
            X.LeMess3 := FloatToStr(Abs(Cha));
            X.LeMess4 := FloatToStr(Abs(Pro));
            ListErr.Add(X);
            Err := not ResultOk;
        end;
    end;
    Ferme(Q);

    { --------- Calcul des comptes sur N-1 pour report sur les comptes géné.*}
    // en fait je ne trouve pas ou est renseigné le SQL de cette query ...
(*
    QLong.Close;
    QLong.Open;
    while not QLong.Eof do
    begin
        Fillchar(Sum, SizeOf(Sum), #0);
        Sum[1].TotDebit := Arrondi(QLong.FindField('D').AsFloat, V_PGI.OkdecV);
        Sum[1].TotCredit := Arrondi(QLong.FindField('C').AsFloat, V_PGI.OkdecV);
        Q := OpenSql(' select sum (g_totdebano) as D, sum(g_totcreano) as C from generaux'
            + ' where g_general="' + QLong.FindField('e_general').AsString + '" and (g_naturegene<>"CHA" and g_naturegene<>"PRO" and g_naturegene<>"EXT")'
            + ' and g_collectif<>"X" ', True);
        Sum[2].TotDebit := Arrondi(Q.FindField('D').AsFloat, V_PGI.OkdecV);
        Sum[2].TotCredit := Arrondi(Q.FindField('C').AsFloat, V_PGI.OkdecV);
        if not Q.Eof then
        begin
            Sol1 := LeSolde(Sum[1].TotDebit, Sum[1].TotCredit);
            Sol2 := LeSolde(Sum[2].TotDebit, Sum[2].TotCredit);
        end
        else
        begin
            Sol1 := 0;
            Sol2 := 0;
        end;

        if Arrondi(LeSolde(Sol1, Sol2), V_PGI.OkdecV) <> 0 then
        begin
            if not CO then
            begin
                X := DelInfo.Create;
                X.LeCod := QLong.FindField('e_general').AsString;
                X.LeLib := QLong.FindField('L').AsString;
                X.LeMess := FloatToStr(Sum[2].TotDebit);
                X.LeMess2 := FloatToStr(Sum[2].TotCredit); //Infos Totaux Comptes
                if Sol1 < 0 then
                begin
                    X.LeMess3 := '';
                    X.LeMess4 := FloatToStr(Abs(Sol1));
                end
                else // Infos Totaux N-1 du Compte
                begin
                    X.LeMess3 := FloatToStr(Abs(Sol1));
                    X.LeMess4 := '';
                end;
                ListErr.Add(X);
                Err := True;
            end
            else
            begin
                if Sol1 < 0 then ExecuteSql(' Update GENERAUX set G_TOTCREANO=' + STRFPOINT(abs(Sol1))
                        + ', G_TOTDEBANO=0 Where g_general="' + QLong.FindField('e_general').AsString + '" ')
                else ExecuteSql(' Update GENERAUX set G_TOTCREANO=0'
                        + ', G_TOTDEBANO=' + STRFPOINT(abs(Sol1)) + ' Where g_general="' + QLong.FindField('e_general').AsString + '" ');
            end;
        end;
        Ferme(Q);
        QLong.Next;
    end;
    QLong.Free;
*)

    { Tiers collectifs }
    QE := OpenSql(' select sum (CU_DEBITAN) as D, sum(CU_CREDITAN) as C, CU_COMPTE1 from CUMULS left join GENERAUX on G_general=CU_COMPTE1  where g_collectif="X" group by CU_COMPTE1 ', True);
    while not QE.Eof do
    begin
        Fillchar(Sum, SizeOf(Sum), #0);
        Sum[1].TotDebit := Arrondi(QE.FindField('D').AsFloat, V_PGI.OkdecV);
        Sum[1].TotCredit := Arrondi(QE.FindField('C').AsFloat, V_PGI.OkdecV);
        Q := OpenSql(' select sum (g_totdebano) as D, sum(g_totcreano) as C, max(g_libelle) as L from generaux where g_general="' + QE.FindField('CU_COMPTE1').AsString + '" and g_collectif="X"  ', True);
        Sum[2].TotDebit := Arrondi(Q.FindField('D').AsFloat, V_PGI.OkdecV);
        Sum[2].TotCredit := Arrondi(Q.FindField('C').AsFloat, V_PGI.OkdecV);
        if not Q.Eof then
        begin
            Sol1 := LeSolde(Sum[1].TotDebit, Sum[1].TotCredit);
            Sol2 := LeSolde(Sum[2].TotDebit, Sum[2].TotCredit);
        end
        else
        begin
            Sol1 := 0;
            Sol2 := 0;
        end;

        if Arrondi(LeSolde(Sol1, Sol2), V_PGI.OkdecV) <> 0 then
        begin
            if not CO then
            begin
                X := DelInfo.Create;
                X.LeCod := QE.FindField('CU_COMPTE1').AsString;
                X.LeLib := Q.FindField('L').AsString;
                X.LeMess := FloatToStr(Sum[2].TotDebit);
                X.LeMess2 := FloatToStr(Sum[2].TotCredit); //Infos Totaux Comptes
                if Sol1 < 0 then
                begin
                    X.LeMess3 := '';
                    X.LeMess4 := FloatToStr(Abs(Sol1));
                end
                else // Infos Totaux N-1 du Compte
                begin
                    X.LeMess3 := FloatToStr(Abs(Sol1));
                    X.LeMess4 := '';
                end;
                ListErr.Add(X);
                Err := True;
            end
            else
            begin
                if Sol1 < 0 then
                    ExecuteSql(' Update GENERAUX set G_TOTCREANO=' + STRFPOINT(abs(Sol1))
                        + ', G_TOTDEBANO=0 Where g_general="' + QE.FindField('CU_COMPTE1').AsString + '" ')
                else
                    ExecuteSql(' Update GENERAUX set G_TOTCREANO=0'
                        + ', G_TOTDEBANO=' + STRFPOINT(abs(Sol1)) + ' Where g_general="' + QE.FindField('CU_COMPTE1').AsString + '" ');
            end;
        end;
        Ferme(Q);
        QE.Next;
    end;
    Ferme(QE);
(*********   Partie II   ******************)
    if (not Err) and (not Co) then
    begin
        Q := OpenSql('Select G_GENERAL, G_LIBELLE, G_TOTDEBANO, G_TOTCREANO  from GENERAUX order by G_GENERAL ', True);
        while not Q.Eof do
        begin
            QE := OpenSql(' Select sum(E_DEBIT) as deb, sum(E_CREDIT) as cre from ECRITURE Where e_journal="ANO" '
                + ' and E_GENERAL="' + Q.Findfield('G_GENERAL').AsString + '" and e_exercice="' + VH^.EnCours.Code + '" group by E_JOURNAL, E_EXERCICE', True);
            if (Arrondi(Q.FindField('G_TOTDEBANO').AsFloat, V_PGI.OkdecV) <> Arrondi(QE.FindField('deb').AsFloat, V_PGI.OkdecV)) or
                (Arrondi(Q.FindField('G_TOTCREANO').AsFloat, V_PGI.OkdecV) <> Arrondi(QE.FindField('cre').AsFloat, V_PGI.OkdecV)) then
            begin
                X := DelInfo.Create;
                X.LeCod := Q.FindField('G_GENERAL').AsString;
                X.LeLib := Q.FindField('G_LIBELLE').AsString;
                X.LeMess := FloatToStr(Q.FindField('G_TOTDEBANO').AsFloat);
                X.LeMess2 := FloatToStr(Q.FindField('G_TOTCREANO').AsFloat);
                X.LeMess3 := FloatToStr(QE.FindField('deb').AsFloat);
                X.LeMess4 := FloatToStr(QE.FindField('cre').AsFloat);
                ListErr.Add(X);
                Err := True;
            end;
            Ferme(QE);
            Q.Next;
        end;
        Ferme(Q);
        Q := OpenSql('Select T_AUXILIAIRE, T_LIBELLE, T_TOTDEBANO, T_TOTCREANO  from TIERS order by T_AUXILIAIRE ', True);
        while not Q.Eof do
        begin
            QE := OpenSql(' Select sum(E_DEBIT) as deb, sum(E_CREDIT) as cre from ECRITURE Where e_journal="ANO" '
                + ' and E_AUXILIAIRE="' + Q.Findfield('T_AUXILIAIRE').AsString + '" and e_exercice="' + VH^.EnCours.Code + '" group by E_JOURNAL, E_EXERCICE', True);
            if (Arrondi(Q.FindField('T_TOTDEBANO').AsFloat, V_PGI.OkdecV) <> Arrondi(QE.FindField('deb').AsFloat, V_PGI.OkdecV)) or
                (Arrondi(Q.FindField('T_TOTCREANO').AsFloat, V_PGI.OkdecV) <> Arrondi(QE.FindField('cre').AsFloat, V_PGI.OkdecV)) then
            begin
                X := DelInfo.Create;
                X.LeCod := Q.FindField('T_AUXILIAIRE').AsString;
                X.LeLib := Q.FindField('T_LIBELLE').AsString;
                X.LeMess := FloatToStr(Q.FindField('T_TOTDEBANO').AsFloat);
                X.LeMess2 := FloatToStr(Q.FindField('T_TOTCREANO').AsFloat);
                X.LeMess3 := FloatToStr(QE.FindField('deb').AsFloat);
                X.LeMess4 := FloatToStr(QE.FindField('cre').AsFloat);
                ListErr.Add(X);
                Err := True;
            end;
            Ferme(QE);
            Q.Next;
        end;
        Ferme(Q);
        Q := OpenSql('Select S_AXE, S_SECTION, S_LIBELLE, S_TOTDEBANO, S_TOTCREANO  from SECTION order by S_AXE, S_SECTION ', True);
        while not Q.Eof do
        begin
            QE := OpenSql(' Select sum(Y_DEBIT) as deb, sum(Y_CREDIT) as cre from ANALYTIQ Where y_journal="ANO" '
                + ' and y_axe="' + Q.Findfield('S_AXE').AsString + '" '
                + ' and y_section="' + Q.FindField('s_section').AsString + '" and y_exercice="' + VH^.EnCours.Code + '"', True);
            if (Arrondi(Q.FindField('S_TOTDEBANO').AsFloat, V_PGI.OkdecV) <> Arrondi(QE.FindField('deb').AsFloat, V_PGI.OkdecV)) or
                (Arrondi(Q.FindField('S_TOTCREANO').AsFloat, V_PGI.OkdecV) <> Arrondi(QE.FindField('cre').AsFloat, V_PGI.OkdecV)) then
            begin
                X := DelInfo.Create;
                X.LeCod := Q.FindField('S_AXE').AsString + ', ' + Q.FindField('S_SECTION').AsString;
                X.LeLib := Q.FindField('S_LIBELLE').AsString;
                X.LeMess := FloatToStr(Q.FindField('S_TOTDEBANO').AsFloat);
                X.LeMess2 := FloatToStr(Q.FindField('S_TOTCREANO').AsFloat);
                X.LeMess3 := FloatToStr(QE.FindField('deb').AsFloat);
                X.LeMess4 := FloatToStr(QE.FindField('cre').AsFloat);
                ListErr.Add(X);
                Err := True;
            end;
            Ferme(QE);
            Q.Next;
        end;
        Ferme(Q);
    end;
    if (Err and (not CO)) then RapportdeSuppression(ListErr, 5);
    ListErr.Clear;
    ListErr.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function LeSolde(D, C : Double) : Double;
begin
    if ((D = C) or (Abs(D) = Abs(C))) then result := 0
    else
        if (Abs(D) > Abs(C)) then result := abs(D) - abs(C)
    else
        if (Abs(D) < Abs(C)) then result := (abs(C) - abs(D)) * -1
    else result := 0;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function VirguleToPoint(Montant : Double) : string;
var
    StMontant : string;
begin
    StMontant := FloatToStr(Montant);
    while (Pos(',', StMontant) > 0) do StMontant[Pos(',', StMontant)] := '.';
    Result := StMontant;
end;

//==================================================
// autres fonctions hors class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure MvtToIdent(var IP : tIdentPiece;Mvt : tInfoMvt);
begin
    FillChar(IP, SizeOf(IP), #0);
    with MVt do
    begin
        IP.JalP := Journal;
        IP.NumP := NumeroPiece;
        IP.DateP := DateComptable;
        IP.QualP := QualifPiece;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function RupturePiece(var IP : tIdentPiece;Mvt : tInfoMvt) : Boolean;
begin
    with Mvt do
        Result := (IP.NumP <> NumeroPiece) or (IP.QualP <> QualifPiece) or (IP.JalP <> JOURNAL);
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure retoucheDelInfo(X : DelInfo;Entitee : TInfoMvt);
begin
    if PourNewImport then
    begin
        X.Jal := Entitee.Journal;
        X.Qualif := Entitee.QualifPiece;
        X.NumP := Entitee.NumeroPiece;
        X.Chrono := Entitee.Chrono;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function ClePieceFausse(X : DelInfo) : string;
begin
    //Result := Format_String(X.Jal,3) + Format_String(X.Qualif,1) + FormatFloat('00000000',X.NumP);
    Result := Format_String(X.Jal, 3) + Format_String(' ', 1) + FormatFloat('00000000', X.NumP);
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function ClePieceAnaFausse(Mvt : TInfoMvt) : string;
begin
    Result := Mvt.Journal + ';' + DateToStr(Mvt.DateComptable) + ';' + Mvt.General + ';' + Mvt.RefInterne + ';';
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure RecupAxe(St : string;var SetAxe : tSetAxe);
var
    IAxeEnCours : Byte;
begin
    IAxeEnCours := StrToInt(Copy(St, 2, 1));
    if IAxeEnCours in [1..5] then SetAxe := SetAxe + [IAxeEnCours];
end;

//==================================================
// Evenements de la fiche
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.FormCreate(Sender : TObject);
begin
    if (not PourNewImport) then LaListe := TList.Create;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.PrepareShow;
begin
    PopUpMenu := ADDMenuPop(PopUpMenu, '', '');
    if Maj then FExercice.DataType := 'ttExoSaufPrecedent';

    PourImport := False;
    FJal1.Text := '';
    FJal2.Text := '';
    DejaLance := False;
    TFJalJoker.Visible := False;
    FJalJoker.Visible := False;
    FEtab.ItemIndex := 0;
    TTemp.Caption := '';
    FSolde.Visible := False;

// GC* - 19/02/2002
// Si en Serie et Verification des mouvements Comptables
    if (VH^.EnSerie) and (FVerification.ItemIndex = 1) then
    begin
        if CtxPCL in V_Pgi.PGIContexte then FExercice.Value := VH^.CPExoRef.Code
        else FExercice.Value := VH^.Encours.Code;

        CtrlDetGen.Checked := True;
        CtrlDetAna.Checked := True;
        CtrlEcrGen.Checked := False;
        CtrlEcrAna.Checked := False;
        CtrlAnaOff.Checked := True;
    end;
// Fin - GC*

    FTypeEcriture.ItemIndex := 0;
    FVerificationChange(nil);
    InitLabel;
    MAJ := False;
    if PourImport then
    begin
        FE := 'IE';
        FY := FE;
        FBE := FE;
        FICECR := 'IMPECR';
        FICANA := FICECR;
        FICBUDECR := FICECR;
    end
    else
    begin
        FE := 'E';
        FY := 'Y';
        FBE := 'BE';
        FICECR := 'ECRITURE';
        FICANA := 'ANALYTIQ';
        FICBUDECR := 'BUDECR';
//        if V_PGI.SAV Then
//        begin
        CtrlDetGen.Enabled := TRUE;
        CtrlDetAna.Enabled := TRUE;
        CtrlEcrGen.Enabled := TRUE;
        CtrlEcrAna.Enabled := TRUE;
//        end;
    end;
{$IFDEF REPARANA}
    CtrlDetGen.Checked := FALSE;
    CtrlDetAna.Checked := FALSE;
    CtrlEcrGen.Checked := FALSE;
    CtrlEcrAna.Checked := FALSE;
    CtrlAnaOff.Checked := TRUE;
    FVerification.ItemIndex := 1;
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.FormShow(Sender : TObject);
begin
    PrepareShow;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.FormClose(Sender : TObject;var Action : TCloseAction);
begin
    if (not PourNewImport) then LaListe.Free;
    {JP 04/08/05 : Cela évite des messages du genre 9 requêtes non fermées en SAV}
    Ferme(QSum);
    Ferme(QEcr);
    Ferme(QAnal);
    Ferme(QInfoAnaG);
    Ferme(QLettrage);
    Ferme(QEquilLet);
    Ferme(QLetDev);
    Ferme(QAnaGene);
    Ferme(QSumGeAn);
    Ferme(QEcrBud);

    FreeAndNil(QSum);
    FreeAndNil(QEcr);
    FreeAndNil(QAnal);
    FreeAndNil(QInfoAnaG);
    FreeAndNil(QLettrage);
    FreeAndNil(QEquilLet);
    FreeAndNil(QLetDev);
    FreeAndNil(QAnaGene);
    FreeAndNil(QSumGeAn);
    FreeAndNil(QEcrBud);
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.BValiderClick(Sender : TObject);
begin
    //StartC(1);
    OkVerif := True;
    MAJ := False;
    NbError := 0;
    ErrEcr := 0;
    ErrMvtA := 0;
    ErrMvtB := 0;
    ErrAnal := 0;
    ErrMvt := 0;
    ErrLet := 0;
    StopVerif := False;

    if DateOk then
    begin
        LaListe.Clear;
        TNBError1.Caption := MsgBar.Mess[14];
        EnableControls(Self, False);
        TTravail.Enabled := True;
        Application.ProcessMessages;
        if SiEnreg then
        begin
            DejaLance := True;
            ChargeInfos;
            LanceVerif;
            if CtrlAnaOff.Checked then lanceVerifAnaOff;
            TTemp.Caption := '';
            //StopC(1) ;ShowC(1,'Total Temp ', False) ;
            if not OkVerif then
            begin
                RapportdErreurMvt(Laliste, 3, MAJ, False);
            end
            else
                if not StopVerif then
                if not VH^.EnSerie then MsgRien.Execute(3, '', ''); // sinon message tout est ok
            VideInfos;
        end
        else
            if not VH^.EnSerie then MsgRien.Execute(0, '', ''); // message : aucun enreg

        InitLabel;
        //If MAJ then Reparation Else Bouton(True) ;
        if MAJ then
        begin
            Reparation;
            InitLabel;
        end;
        EnableControls(Self, True);
    end
    else MsgRien.Execute(2, '', '');

    if FVerification.ItemIndex <> 1 then
        if LetListe <> nil then VideLetListe; {Vide Infos lettragDev incorrect}
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.BStopClick(Sender : TObject);
begin
    StopVerif := True;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.FJal1Change(Sender : TObject);
var
    AvecJoker : Boolean;
begin
    AvecJoker := Joker(FJal1, FJal2, FJalJoker);
    TFaJ.Visible := not AvecJoker;
    TFJal.Visible := not AvecJoker;
    TFJalJoker.Visible := AvecJoker;

    if not AvecJoker then
        if not VH^.EnSerie then
            if THCpteEdit(Sender).CanFocus then THCpteEdit(Sender).SetFocus;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.FExerciceChange(Sender : TObject);
begin
(* CR le 18/12/97
    if ((FExercice.Value<>V_PGI.Precedent.Code) and (V_PGI.Precedent.Code<>'')) then ExoToDates(FExercice.Value,FDateCpta1,FDateCpta2)
    else if FExercice.Value<>V_PGI.Encours.Code then FExercice.Value:=V_PGI.Encours.Code;
*)
    ExoToDates(FExercice.Value, FDateCpta1, FDateCpta2);
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFVerCpta.FVerificationChange(Sender : TObject);
begin
    if FVerification.ItemIndex = 2 then
    begin
        FExercice.Vide := True;
        FExercice.ItemIndex := -1; // Buggé à mort !!!
        FExercice.Text := MsgBar.Mess[19];

        if FJal1.Enabled then
            if FJal1.Visible then FJal1.Clear;
        if FJal2.Enabled then
            if FJal2.Visible then FJal2.Clear;
        if FJalJoker.Enabled then
            if FJalJoker.Visible then FJalJoker.Clear;
        if VH^.Precedent.Code <> '' then FDateCpta1.Text := DateToStr(VH^.Precedent.Deb)
        else FDateCpta1.Text := DateToStr(VH^.Encours.Deb);
        if VH^.Suivant.Code <> '' then FDateCpta2.Text := DateToStr(VH^.Suivant.Fin)
        else FDateCpta2.Text := DateToStr(VH^.Encours.Fin);
    end
    else
    begin
        FExercice.Vide := False;
        FExercice.Value := VH^.Encours.Code;
    end;
    if FVerification.ItemIndex = 2 then
    begin
        TFJal.Enabled := (FVerification.ItemIndex <> 2);
        FJal1.Enabled := (FVerification.ItemIndex <> 2);
        FJal2.Enabled := (FVerification.ItemIndex <> 2);
        TFaJ.Enabled := (FVerification.ItemIndex <> 2);
        TFJalJoker.Enabled := (FVerification.ItemIndex <> 2);
        FJalJoker.Enabled := (FVerification.ItemIndex <> 2);
        TFExercice.Enabled := (FVerification.ItemIndex <> 2);
        FExercice.Enabled := (FVerification.ItemIndex <> 2);
        TFDateCpta1.Enabled := (FVerification.ItemIndex <> 2);
        TFDateCpta2.Enabled := (FVerification.ItemIndex <> 2);
        FDateCpta1.Enabled := (FVerification.ItemIndex <> 2);
        FDateCpta2.Enabled := (FVerification.ItemIndex <> 2);
        TFTypeEcriture.Enabled := (FVerification.ItemIndex <> 2);
        FTypeEcriture.Enabled := (FVerification.ItemIndex <> 2);
        TFNumPiece1.Enabled := (FVerification.ItemIndex <> 2);
        FNumPiece1.Enabled := (FVerification.ItemIndex <> 2);
        FNumPiece2.Enabled := (FVerification.ItemIndex <> 2);
        TFNumPiece2.Enabled := (FVerification.ItemIndex <> 2);
    end;
    if FVerification.ItemIndex = 3 then
    begin
        FJal1.ZoomTable := tzBudJal;
        FJal2.ZoomTable := tzBudJal;
        FExercice.DataType := 'ttExerciceBudget';
        // Encours aussi budgétaire ou suivant ?
        FExercice.Value := VH^.Encours.Code;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.FDateCpta1KeyPress(Sender : TObject;var Key : Char);
begin
    ParamDate(Self, Sender, Key);
end;

//==================================================
// Autres fonctions de la class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.SiEnreg : Boolean;
var
    AuMoinsUn : Boolean;
    sql : string;
begin
    Result := False;
    NbEnreg := 0;
    AuMoinsUn := FALSE;

    //If LanceDirect then begin Result:=True ;InitMove(1000,MsgBar.Mess[0]) ; Exit ; end ;

    if (FVerification.ItemIndex = 2) then {que la verif Lettrage}
    begin
        AuMoinsUn := CompteLettrage;
        if (AuMoinsUn) then
        begin
            InitMove(1000, MsgBar.Mess[0]);
            Result := True;
            Exit;
        end;
    end;

{----------------------------------------------------------------------------------------------}

    if CtrlEcrGen.Checked then
    begin
        { Somme D et C (Devise pivot et devise) de chaque Pièce }

        if (not (QSum = nil)) then Ferme(QSum);

        if (PourImport)
          then SQL := 'SELECT IE_JOURNAL, IE_LETTREPOINTLCR, IE_NUMPIECE,     IE_DEVISE,'
          else SQL := 'SELECT E_JOURNAL,  E_EXERCICE,        E_DATECOMPTABLE, E_NUMEROPIECE, E_DEVISE,';

        SQL := SQL + ' SUM(' + FE + '_DEBIT) AS SUMD, SUM(' + FE + '_CREDIT) AS SUMC, SUM(' + FE + '_DEBITDEV) AS SumDDev, Sum(' + FE + '_CREDITDEV) as SumCDev, ';
        SQL := SQL + ' MIN(' + FE + '_DEBIT) AS MinD, MAX(' + FE + '_DEBIT) as MaxD,  MIN(' + FE + '_CREDIT) AS MinC,      MAX(' + FE + '_CREDIT) as MaxC ';

        if (PourImport)
          then SQL := SQL + ', MAX(' + FE + '_CONTROLE) as MaxCONTROLE '
          else SQL := SQL + ', E_NATUREPIECE ' ;         // FQ17137

        SQL := SQL + ' FROM ' + FICECR + ' WHERE ' + FE + '_JOURNAL<>"' + w_w + '" ';

        if (not PourImport)
          then SQL := SQL + ' And E_EXERCICE="' + VCrit.Exo.Code + '" AND (E_MODESAISIE="-" OR E_MODESAISIE="") ';
        if (VCrit.QualifPiece <> '')
          then SQL := SQL + ' And ' + FE + '_QUALIFPIECE="' + VCrit.QualifPiece + '" ';
        if (VCrit.Joker)
          then SQL := SQL + ' And ' + FE + '_JOURNAL like "' + TraduitJoker(VCrit.Cpt1) + '"'
          else
            begin
            if (VCrit.Cpt1 <> '') then SQL := SQL + ' And ' + FE + '_JOURNAL>="' + VCrit.Cpt1 + '"';
            if (VCrit.Cpt2 <> '') then SQL := SQL + ' And ' + FE + '_JOURNAL<="' + VCrit.Cpt2 + '"';
            end;

        if (not PourImport)
          then SQL := SQL + ' And E_NUMEROPIECE>=' + IntToStr(VCrit.GL.NumPiece1) + ' and E_NUMEROPIECE<=' + IntTostr(VCrit.GL.NumPiece2) + ' '
          else SQL := SQL + ' And IE_NUMPIECE>=' + IntToStr(VCrit.GL.NumPiece1) + ' and IE_NUMPIECE<=' + IntTostr(VCrit.GL.NumPiece2) + ' ';

        if (VCrit.Etab <> '')
          then SQL := SQL + ' And ' + FE + '_ETABLISSEMENT="' + VCrit.Etab + '" ';
        if (PourImport)
          then SQL := SQL + ' And IE_TYPEECR="E" and IE_SELECTED="X" and IE_OKCONTROLE<>"D" ';
        SQL := SQL + ' and ' + FE + '_QUALIFPIECE<>"C" ';

        if (PourImport)
          then SQL := SQL + ' GROUP BY IE_JOURNAL, IE_LETTREPOINTLCR, IE_NUMPIECE, IE_DEVISE '
          else SQL := SQL + ' GROUP BY E_JOURNAL,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_DEVISE,E_NATUREPIECE Having ' + FE + '_DATECOMPTABLE>="' + USDATETIME(VCrit.Date1) + '" And ' + FE + '_DATECOMPTABLE<="' + USDATETIME(VCrit.Date2) + '"';

        QSum := OpenSQl(SQL,true);

        AuMoinsUn := AuMoinsUn or (not QSum.EOF);
    end;
    Application.ProcessMessages;

{--------------------------------------------------------------------------------------------}

    if CtrlEcrAna.Checked then
    begin
        { Somme D et C (Devise pivot et devise) de chaque Pièce/Ligne }

        if (not (QSumGeAn = nil)) then Ferme(QSumGeAn);

        if (PourImport) then SQL := ' Select IE_JOURNAL, IE_DATECOMPTABLE, IE_NUMPIECE, IE_DEVISE, '
        else SQL := ' Select E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_DEVISE, ';
        SQL := SQL + ' ' + FE + '_DEBIT, ' + FE + '_CREDIT, ' + FE + '_DEBITDEV, ' + FE + '_CREDITDEV, ';
        SQL := SQL + ' ' + FE + '_NUMLIGNE, ';
        SQL := SQL + ' ' + FE + '_REFINTERNE, ' + FE + '_GENERAL, ' + FE + '_ETABLISSEMENT, ';
        SQL := SQL + ' ' + FE + '_NATUREPIECE, ' + FE + '_ECRANOUVEAU,  ';
        SQL := SQL + ' G_VENTILABLE1, G_VENTILABLE2, G_VENTILABLE3, G_VENTILABLE4, G_VENTILABLE5, ';
        SQL := SQL + ' ' + FE + '_TYPEANOUVEAU, ' + FE + '_QUALIFPIECE ';
        if (PourImport) then SQL := SQL + ' ,IE_CHRONO ';
        SQL := SQL + ' From ' + FICECR + ' ';
        SQL := SQL + ' Left Join GENERAUX on ' + FE + '_GENERAL=G_GENERAL where ' + FE + '_ANA="X" ';
        if (not PourImport) then SQL := SQL + ' And E_EXERCICE="' + VCrit.Exo.Code + '" ';
        if (VCrit.QualifPiece <> '') then SQL := SQL + ' And ' + FE + '_QUALIFPIECE="' + VCrit.QualifPiece + '" ';
        if (VCrit.Joker) then SQL := SQL + ' And ' + FE + '_JOURNAL like "' + TraduitJoker(VCrit.Cpt1) + '"'
        else
        begin
            if (VCrit.Cpt1 <> '') then SQL := SQL + ' And ' + FE + '_JOURNAL>="' + VCrit.Cpt1 + '"';
            if (VCrit.Cpt2 <> '') then SQL := SQL + ' And ' + FE + '_JOURNAL<="' + VCrit.Cpt2 + '"';
        end;
        if (not PourImport) then SQL := SQL + ' And E_NUMEROPIECE>=' + IntToStr(VCrit.GL.NumPiece1) + ' and E_NUMEROPIECE<=' + IntTostr(VCrit.GL.NumPiece2) + ' '
        else SQL := SQL + ' And IE_NUMPIECE>=' + IntToStr(VCrit.GL.NumPiece1) + ' and IE_NUMPIECE<=' + IntTostr(VCrit.GL.NumPiece2) + ' ';
        SQL := SQL + ' And ' + FE + '_DATECOMPTABLE>="' + USDATETIME(VCrit.Date1) + '" ';
        SQL := SQL + ' And ' + FE + '_DATECOMPTABLE<="' + USDATETIME(VCrit.Date2) + '" ';
        if (VCrit.Etab <> '') then SQL := SQL + ' And ' + FE + '_ETABLISSEMENT="' + VCrit.Etab + '" ';
        if (PourImport) then SQL := SQL + ' And IE_TYPEECR="E" and IE_SELECTED="X" and IE_OKCONTROLE<>"D" ';
        SQL := SQL + ' And ' + FE + '_QUALIFPIECE<>"C" ';
        { Construction de la clause Order By de la SQL }
        if (PourImport) then SQL := SQL + ' Order BY IE_CHRONO '
        else SQL := SQL + ' Order BY E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_DEVISE, E_QUALIFPIECE ';

        QSumGeAn := OpenSQL(SQL,true);

        AuMoinsUn := AuMoinsUn or (not QSumGeAn.EOF);
    end;
    Application.ProcessMessages;

{----------------------------------------------------------------------------------------------}

{ Verif des Mvts Analytique }
    if CtrlDetAna.Checked then
    begin
        if (not (QAnal = nil)) then Ferme(QAnal);

        SQL := ' Select ' + FY + '_JOURNAL, ' + FY + '_DATECOMPTABLE, ' + FY + '_NUMLIGNE,  ';
        SQL := SQL + ' ' + FY + '_AXE, ' + FY + '_DEVISE, ' + FY + '_REFINTERNE, ' + FY + '_GENERAL, ' + FY + '_NUMVENTIL, ';
        SQL := SQL + ' ' + FY + '_SECTION, ' + FY + '_ECRANOUVEAU,  ' + FY + '_NATUREPIECE, ' + FY + '_ETABLISSEMENT, ';
        SQL := SQL + ' ' + FY + '_TYPEANALYTIQUE, ';
        SQL := SQL + ' ' + FY + '_DEBIT, ' + FY + '_CREDIT, ' + FY + '_DEBITDEV, ' + FY + '_CREDITDEV, ';
        if (not PourImport) then SQL := SQL + ' Y_NUMEROPIECE, Y_CONFIDENTIEL, Y_EXERCICE, '
        else SQL := SQL + ' IE_NUMPIECE, IE_CHRONO, ';
        SQL := SQL + ' ' + FY + '_TAUXDEV, ' + FY + '_QUALIFPIECE ';
        SQL := SQL + ' From ' + FICANA + ' ';
        SQL := SQL + ' Where ';
        SQL := SQL + ' ' + FY + '_JOURNAL<>"' + w_w + '" ';
        if (not PourImport) then SQL := SQL + ' And Y_EXERCICE="' + VCrit.Exo.Code + '" ';
        if (VCrit.QualifPiece <> '') then SQL := SQL + ' And ' + FY + '_QUALIFPIECE="' + VCrit.QualifPiece + '" ';
        if (VCrit.Joker) then SQL := SQL + ' And ' + FY + '_JOURNAL like "' + TraduitJoker(VCrit.Cpt1) + '" '
        else
        begin
            if (VCrit.Cpt1 <> '') then SQL := SQL + ' And ' + FY + '_JOURNAL>="' + VCrit.Cpt1 + '" ';
            if (VCrit.Cpt2 <> '') then SQL := SQL + ' And ' + FY + '_JOURNAL<="' + VCrit.Cpt2 + '" ';
        end;
        if (not PourImport) then SQL := SQL + ' And Y_NUMEROPIECE>=' + IntToStr(VCrit.GL.NumPiece1) + ' and Y_NUMEROPIECE<=' + IntTostr(VCrit.GL.NumPiece2) + ' '
        else SQL := SQL + ' And IE_NUMPIECE>=' + IntToStr(VCrit.GL.NumPiece1) + ' and IE_NUMPIECE<=' + IntTostr(VCrit.GL.NumPiece2) + ' ';
        SQL := SQL + ' And ' + FY + '_DATECOMPTABLE>="' + USDATETIME(VCrit.Date1) + '" ';
        SQL := SQL + ' And ' + FY + '_DATECOMPTABLE<="' + USDATETIME(VCrit.Date2) + '" ';
        if (VCrit.Etab <> '') then SQL := SQL + ' And ' + FY + '_ETABLISSEMENT="' + VCrit.Etab + '" ';
        if (PourImport) then SQL := SQL + ' And IE_TYPEECR="A" and IE_SELECTED="X" and IE_OKCONTROLE<>"D" ';
        SQL := SQL + ' and ' + FY + '_QUALIFPIECE<>"C" ';
        { Construction de la clause Order By de la SQL }
        if (not PourImport) then SQL := SQL + ' Order By Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_AXE, Y_NUMVENTIL,Y_QUALIFPIECE '
        else SQL := SQL + ' Order By IE_CHRONO ';

        QAnal := OpenSQL(SQL,true);

        AuMoinsUn := AuMoinsUn or (not QAnal.EOF);
    end;
    Application.ProcessMessages;

{----------------------------------------------------------------------------------------------}

    if CtrlDetGen.Checked then
    begin
        if (not (QEcr = nil)) then Ferme(QEcr);

        SQL := ' Select ' + FE + '_JOURNAL, ' + FE + '_DATECOMPTABLE, ' + FE + '_NUMLIGNE,  ';
        SQL := SQL + ' ' + FE + '_GENERAL, ' + FE + '_AUXILIAIRE,  ' + FE + '_DEVISE, ' + FE + '_MODEPAIE, ';
        SQL := SQL + ' ' + FE + '_DEBIT,  ' + FE + '_CREDIT , ' + FE + '_DEBITDEV,  ' + FE + '_CREDITDEV, ';
        SQL := SQL + ' ' + FE + '_ANA, ' + FE + '_REFINTERNE, ' + FE + '_DATEECHEANCE, ';
        SQL := SQL + ' ' + FE + '_NUMECHE, ' + FE + '_ETABLISSEMENT, ' + FE + '_ETATLETTRAGE, ' + FE + '_ECHE, ';
        SQL := SQL + ' ' + FE + '_ECRANOUVEAU, ' + FE + '_COUVERTURE, ' + FE + '_COUVERTUREDEV, ' + FE + '_NATUREPIECE, ';
        if not PourImport then SQL := SQL + ' E_NUMEROPIECE, E_CONFIDENTIEL, E_EXERCICE,  '
        else SQL := SQL + ' IE_NUMPIECE, IE_CHRONO, ';
        SQL := SQL + ' ' + FE + '_TAUXDEV, ' + FE + '_QUALIFPIECE ';
        if not PourImport then SQL := SQL + ', ' + FE + '_MODESAISIE '
        else SQL := SQL + ', ' + FE + '_LIBREBOOL1 ';
        SQL := SQL + ' From ' + FICECR + ' ';
        SQL := SQL + ' Where ' + FE + '_JOURNAL<>"' + w_w + '" ';
        if not PourImport then SQL := SQL + ' And E_EXERCICE="' + VCrit.Exo.Code + '" ';
        if VCrit.QualifPiece <> '' then SQL := SQL + ' And ' + FE + '_QUALIFPIECE="' + VCrit.QualifPiece + '" ';
        if VCrit.Joker then SQL := SQL + ' And ' + FE + '_JOURNAL like "' + TraduitJoker(VCrit.Cpt1) + '"'
        else
        begin
            if VCrit.Cpt1 <> '' then SQL := SQL + ' And ' + FE + '_JOURNAL>="' + VCrit.Cpt1 + '"';
            if VCrit.Cpt2 <> '' then SQL := SQL + ' And ' + FE + '_JOURNAL<="' + VCrit.Cpt2 + '"';
        end;
        if not PourImport then SQL := SQL + ' And E_NUMEROPIECE>=' + IntToStr(VCrit.GL.NumPiece1) + ' and E_NUMEROPIECE<=' + IntTostr(VCrit.GL.NumPiece2) + ' '
        else SQL := SQL + ' And IE_NUMPIECE>=' + IntToStr(VCrit.GL.NumPiece1) + ' and IE_NUMPIECE<=' + IntTostr(VCrit.GL.NumPiece2) + ' ';
        SQL := SQL + ' And ' + FE + '_DATECOMPTABLE>="' + USDATETIME(VCrit.Date1) + '"';
        SQL := SQL + ' And ' + FE + '_DATECOMPTABLE<="' + USDATETIME(VCrit.Date2) + '"';
        if VCrit.Etab <> '' then SQL := SQL + ' And ' + FE + '_ETABLISSEMENT="' + VCrit.Etab + '" ';
        if PourImport then SQL := SQL + ' And IE_TYPEECR="E" and IE_SELECTED="X" and IE_OKCONTROLE<>"D" ';
        { Construction de la clause Order By de la SQL }
        SQL := SQL + ' And ' + FE + '_QUALIFPIECE<>"C" ';
        if not PourImport then SQL := SQL + ' order BY E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE '
        else SQL := SQL + ' order BY IE_CHRONO ';

        QEcr := OpenSQL(SQL,true);

        AuMoinsUn := AuMoinsUn or (not QEcr.EOF);
    end;
    Application.ProcessMessages;

{----------------------------------------------------------------------------------------------}

{ Vérification mouvements budgétaires }
    if (FVerification.ItemIndex = 3) or (FVerification.ItemIndex = 0) then
    begin
        if (not (QEcrBud = nil)) then Ferme(QEcrBud);

        SQL := 'Select ' + FBE + '_DATECOMPTABLE, ';
        SQL := SQL + ' ' + FBE + '_DEBIT,  ' + FBE + '_CREDIT , ' + FBE + '_AXE, ';
        SQL := SQL + ' ' + FBE + '_REFINTERNE, ' + FBE + '_ETABLISSEMENT, ';
        if not PourImport then SQL := SQL + ' BE_EXERCICE, BE_NUMEROPIECE, BE_CONFIDENTIEL, BE_BUDJAL, BE_NATUREBUD, BE_BUDGENE, BE_BUDSECT, '
        else SQL := SQL + ' IE_NUMPIECE, IE_CHRONO, IE_JOURNAL, IE_NATUREPIECE, IE_GENERAL, IE_SECTION, ';
        SQL := SQL + ' ' + FBE + '_QUALIFPIECE ';
        SQL := SQL + ' From ' + FICBUDECR + ' ';
        if not PourImport then SQL := SQL + ' Where ' + FBE + '_BUDJAL<>"' + w_w + '" '
        else SQL := SQL + ' Where ' + FBE + '_JOURNAL<>"' + w_w + '" ';
        if not PourImport then SQL := SQL + ' And BE_EXERCICE="' + VCrit.Exo.Code + '" ';
        if VCrit.QualifPiece <> '' then SQL := SQL + ' And ' + FBE + '_QUALIFPIECE="' + VCrit.QualifPiece + '" ';
        if VCrit.Joker then
        begin
            if not PourImport then SQL := SQL + ' And ' + FBE + '_BUDJAL like "' + TraduitJoker(VCrit.Cpt1) + '"'
            else SQL := SQL + ' And ' + FBE + '_JOURNAL like "' + TraduitJoker(VCrit.Cpt1) + '"';
        end
        else
        begin
            if not PourImport then
            begin
                if VCrit.Cpt1 <> '' then SQL := SQL + ' And ' + FBE + '_BUDJAL>="' + VCrit.Cpt1 + '"';
                if VCrit.Cpt2 <> '' then SQL := SQL + ' And ' + FBE + '_BUDJAL<="' + VCrit.Cpt2 + '"';
            end
            else
            begin
                if VCrit.Cpt1 <> '' then SQL := SQL + ' And ' + FBE + '_JOURNAL>="' + VCrit.Cpt1 + '"';
                if VCrit.Cpt2 <> '' then SQL := SQL + ' And ' + FBE + '_JOURNAL<="' + VCrit.Cpt2 + '"';
            end;
        end;
        if not PourImport then SQL := SQL + ' And BE_NUMEROPIECE>=' + IntToStr(VCrit.GL.NumPiece1) + ' and BE_NUMEROPIECE<=' + IntTostr(VCrit.GL.NumPiece2) + ' '
        else SQL := SQL + ' And IE_NUMPIECE>=' + IntToStr(VCrit.GL.NumPiece1) + ' and IE_NUMPIECE<=' + IntTostr(VCrit.GL.NumPiece2) + ' ';
        SQL := SQL + ' And ' + FBE + '_DATECOMPTABLE>="' + USDATETIME(VCrit.Date1) + '"';
        SQL := SQL + ' And ' + FBE + '_DATECOMPTABLE<="' + USDATETIME(VCrit.Date2) + '"';
        if VCrit.Etab <> '' then SQL := SQL + ' And ' + FBE + '_ETABLISSEMENT="' + VCrit.Etab + '" ';
        if PourImport then SQL := SQL + ' And IE_TYPEECR="E" and IE_SELECTED="X" and IE_OKCONTROLE<>"D" ';
        { Construction de la clause Order By de la SQL }
        SQL := SQL + ' And ' + FBE + '_QUALIFPIECE<>"C" ';
        if not PourImport then SQL := SQL + ' order BY BE_BUDJAL, BE_EXERCICE, BE_DATECOMPTABLE, BE_QUALIFPIECE '
        else SQL := SQL + ' order BY IE_CHRONO ';

        QEcrBud := OpenSQL(SQl,true);

        AuMoinsUn := AuMoinsUn or (not QEcrBud.EOF);
    end;

{----------------------------------------------------------------------------------------------}

    if FVerification.ItemIndex = 0 then AuMoinsUn := CompteLettrage; {Avec la Verif Lettrage}
    if (not AuMoinsUn) and CtrlAnaOff.Checked then AuMoinsUn := TRUE;

    if AuMoinsUn then
    begin
        Result := True;
        InitMove(10000, MsgBar.Mess[0]);
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.ChargeInfos;
var
    Q : TQuery;
    infos : TEnregInfos;
    NbAux : Integer;
begin
    ListeDev := TStringList.Create;
    Q := OpenSql(' Select D_DEVISE, D_DECIMALE, D_QUOTITE from DEVISE ', True);
    while not Q.Eof do
    begin
        Infos := TEnregInfos.Create;
        Infos.Inf1 := Q.Fields[0].AsString;
        Infos.Nb := Q.Fields[1].AsInteger;
        Infos.Mont := Q.Fields[2].AsFloat;
        ListeDev.AddObject('', Infos);
        Q.Next;
    end;
    Ferme(Q);
    if FVerification.ItemIndex = 2 then exit;
    ListeJal := TStringList.Create;
    ListeExo := TStringList.Create;
    ListeSec := TStringList.Create;
    ListeGen := TStringList.Create;
    ListeBQE := TStringList.Create;
    ListeAUX := TStringList.Create;
    ListeJalBud := TStringList.Create;
    ListeCptBud := TStringList.Create;
    ListeSecBud := TStringList.Create;
    Q := OpenSql(' Select J_JOURNAL, J_NATUREJAL, J_AXE, J_COMPTEINTERDIT, J_MODESAISIE, J_VALIDEEN, J_VALIDEEN1 from JOURNAL order by J_JOURNAL ', True); // FQ 12405
    while not Q.Eof do
    begin
        Infos := TEnregInfos.Create;
        Infos.Inf1 := Q.FindField('J_JOURNAL').AsString;
        Infos.Inf2 := Q.FindField('J_NATUREJAL').AsString;
        Infos.Inf3 := Q.FindField('J_AXE').AsString;
        Infos.Inf4 := Q.FindField('J_COMPTEINTERDIT').AsString;
        Infos.Inf5 := Q.FindField('J_MODESAISIE').AsString;
        Infos.Inf6 := Q.FindField('J_VALIDEEN').AsString; // FQ 12405
        Infos.Inf7 := Q.FindField('J_VALIDEEN1').AsString; // FQ 12405
        ListeJal.AddObject('', Infos);
        Q.Next;
    end;
    Ferme(Q);
    Q := OpenSql(' select EX_EXERCICE, EX_DATEDEBUT, EX_DATEFIN from EXERCICE ', True);
    while not Q.Eof do
    begin
        Infos := TEnregInfos.Create;
        Infos.Inf1 := Q.FindField('EX_EXERCICE').AsString;
        Infos.D1 := Q.FindField('EX_DATEDEBUT').AsDateTime;
        Infos.D2 := Q.FindField('EX_DATEFIN').AsDateTime;
        ListeExo.AddObject('', Infos);
        Q.Next;
    end;
    Ferme(Q);

    Q := OpenSql(' select S_AXE, S_SECTION from SECTION order by S_AXE, S_SECTION ', True);
    while not Q.Eof do
    begin
        Infos := TEnregInfos.Create;
        Infos.Inf1 := Q.FindField('S_AXE').AsString;
        Infos.Inf2 := Q.FindField('S_SECTION').AsString;
        ListeSec.AddObject('', Infos);
        Q.Next;
    end;
    Ferme(Q);
    Q := OpenSql(' select BJ_BUDJAL from BUDJAL order by BJ_BUDJAL ', True);
    while not Q.Eof do
    begin
        Infos := TEnregInfos.Create;
        Infos.Inf1 := Q.FindField('BJ_BUDJAL').AsString;
        ListeJalBud.AddObject('', Infos);
        Q.Next;
    end;
    Ferme(Q);
    Q := OpenSql(' select BG_BUDGENE from BUDGENE order by BG_BUDGENE ', True);
    while not Q.Eof do
    begin
        Infos := TEnregInfos.Create;
        Infos.Inf1 := Q.FindField('BG_BUDGENE').AsString;
        ListeCptBud.AddObject('', Infos);
        Q.Next;
    end;
    Ferme(Q);
    Q := OpenSql(' select BS_AXE, BS_BUDSECT from BUDSECT order by BS_AXE, BS_BUDSECT ', True);
    while not Q.Eof do
    begin
        Infos := TEnregInfos.Create;
        Infos.Inf1 := Q.FindField('BS_AXE').AsString;
        Infos.Inf2 := Q.FindField('BS_BUDSECT').AsString;
        ListeSecBud.AddObject('', Infos);
        Q.Next;
    end;
    Ferme(Q);

// Vérif Budgétaire
    if FVerification.ItemIndex = 3 then exit;
    Q := OpenSql('select G_GENERAL, G_NATUREGENE, G_COLLECTIF, G_LETTRABLE, G_POINTABLE, '
        + 'G_VENTILABLE1, G_VENTILABLE2, G_VENTILABLE3, G_VENTILABLE4, G_VENTILABLE5 '
        + 'from GENERAUX order by G_GENERAL', True);
    while not Q.Eof do
    begin
        Infos := TEnregInfos.Create;
        Infos.Inf1 := Q.FindField('G_GENERAL').AsString;
        Infos.Inf2 := Q.FindField('G_NATUREGENE').AsString;
        Infos.Bol1 := (Q.FindField('G_COLLECTIF').AsString = 'X');
        Infos.Bol2 := (Q.FindField('G_LETTRABLE').AsString = 'X');
        Infos.Bol3 := (Q.FindField('G_POINTABLE').AsString = 'X');
        Infos.SuperBol[1] := (Q.FindField('G_VENTILABLE1').AsString = 'X');
        Infos.SuperBol[2] := (Q.FindField('G_VENTILABLE2').AsString = 'X');
        Infos.SuperBol[3] := (Q.FindField('G_VENTILABLE3').AsString = 'X');
        Infos.SuperBol[4] := (Q.FindField('G_VENTILABLE4').AsString = 'X');
        Infos.SuperBol[5] := (Q.FindField('G_VENTILABLE5').AsString = 'X');
        ListeGen.AddObject('', Infos);
        Q.Next;
    end;
    Ferme(Q);

    Q := OpenSql('select BQ_GENERAL, BQ_DEVISE from BANQUECP WHERE BQ_NODOSSIER="'+V_PGI.NoDossier+'"', True); // 19/10/2006 YMO Multisociétés
    while not Q.Eof do
    begin
        Infos := TEnregInfos.Create;
        Infos.Inf1 := Q.FindField('BQ_GENERAL').AsString;
        Infos.Inf2 := Q.FindField('BQ_DEVISE').AsString;
        ListeBQE.AddObject('', Infos);
        Q.Next;
    end;
    Ferme(Q);

    NbAux := 0;
    Q := OpenSql('select T_AUXILIAIRE, T_NATUREAUXI, T_LETTRABLE from TIERS', True);
    while (not Q.Eof) and (NbAux <= 2000) do
    begin
        Infos := TEnregInfos.Create;
        Infos.Inf1 := Q.FindField('T_AUXILIAIRE').AsString;
        Infos.Inf2 := Q.FindField('T_NATUREAUXI').AsString;
        Infos.Bol1 := (Q.FindField('T_LETTRABLE').AsString = 'X');
        ListeAUX.AddObject('', Infos);
        Inc(NbAux);
        Q.Next;
    end;
    Ferme(Q);

    TCNatPiece.Clear;

    Q := OpenSql(' select CO_CODE from COMMUN where CO_TYPE="NTP" ', True);
    while not Q.Eof do
    begin
        TCNatPiece.Items.Add(Q.FindField('CO_CODE').AsString);
        Q.Next;
    end;
    Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.VideInfos;
var
    i : integer;
    infos : TEnregInfos;
begin
    for i := 0 to ListeDev.Count - 1 do
    begin
        infos := TEnregInfos(ListeDev.Objects[i]);
        infos.Free;
    end;
    if ListeDev <> nil then
    begin
        ListeDev.Clear;
        ListeDev.Free;
    end;

    if FVerification.ItemIndex = 2 then exit;

    for i := 0 to ListeJal.Count - 1 do
    begin
        infos := TEnregInfos(ListeJal.Objects[i]);
        infos.Free;
    end;
    if ListeJal <> nil then
    begin
        ListeJal.Clear;
        ListeJal.Free;
    end;

    for i := 0 to ListeExo.Count - 1 do
    begin
        infos := TEnregInfos(ListeExo.Objects[i]);
        infos.Free;
    end;
    if ListeExo <> nil then
    begin
        ListeExo.Clear;
        ListeExo.Free;
    end;

    for i := 0 to ListeSec.Count - 1 do
    begin
        infos := TEnregInfos(ListeSec.Objects[i]);
        infos.Free;
    end;
    if ListeSec <> nil then
    begin
        ListeSec.Clear;
        ListeSec.Free;
    end;

    for i := 0 to ListeJalBud.Count - 1 do
    begin
        infos := TEnregInfos(ListeJalBud.Objects[i]);
        infos.Free;
    end;
    if ListeJalBud <> nil then
    begin
        ListeJalBud.Clear;
        ListeJalBud.Free;
    end;

    for i := 0 to ListeCptBud.Count - 1 do
    begin
        infos := TEnregInfos(ListeCptBud.Objects[i]);
        infos.Free;
    end;
    if ListeCptBud <> nil then
    begin
        ListeCptBud.Clear;
        ListeCptBud.Free;
    end;

    for i := 0 to ListeSecBud.Count - 1 do
    begin
        infos := TEnregInfos(ListeSecBud.Objects[i]);
        infos.Free;
    end;
    if ListeSecBud <> nil then
    begin
        ListeSecBud.Clear;
        ListeSecBud.Free;
    end;

// Vérif Budgétaire
    if FVerification.ItemIndex = 3 then exit;

    for i := 0 to ListeGen.Count - 1 do
    begin
        infos := TEnregInfos(ListeGen.Objects[i]);
        infos.Free;
    end;
    if ListeGen <> nil then
    begin
        ListeGen.Clear;
        ListeGen.Free;
    end;

    for i := 0 to ListeBQE.Count - 1 do
    begin
        infos := TEnregInfos(ListeBQE.Objects[i]);
        infos.Free;
    end;
    if ListeBQE <> nil then
    begin
        ListeBQE.Clear;
        ListeBQE.Free;
    end;

    for i := 0 to ListeAUX.Count - 1 do
    begin
        infos := TEnregInfos(ListeAUX.Objects[i]);
        infos.Free;
    end;
    if ListeAUX <> nil then
    begin
        ListeAUX.Clear;
        ListeAUX.Free;
    end;
    TCNatPiece.Clear;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.InitLabel;
begin
    //if EnSerie then Exit;
    TTravail.Caption := '';
    TNBError1.Caption := '';
    TNBError2.Caption := '';
    TNBError3.Caption := '';
    TNBError4.Caption := '';
    TNBError5.Caption := '';
    TNBError6.Caption := '';
    TNBError1.Font.color := clNavy;
    TNBError2.Font.color := clNavy;
    TNBError3.Font.color := clNavy;
    TNBError4.Font.color := clNavy;
    TNBError5.Font.color := clNavy;
    TNBError6.Font.color := clNavy;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.DateOk : Boolean;
begin
    Result := CtrlPerExo(StrToDate(FDateCpta1.Text), StrToDate(FDateCpta2.Text));
    if Result then RecupCrit;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.VideLetListe;
var
    ErLetDev : TErrLetDev;
    i : integer;
begin
    for i := 0 to LetListe.Count - 1 do
    begin
        ErLetDev := TErrLetDev(LetListe.Objects[i]);
        ErLetDev.Free;
    end;
    LetListe.Clear;
    LetListe.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.CompteLettrage : Boolean;
var
    Rien : Boolean;
    SQL : string;
begin
    if (not (QLettrage = nil)) then Ferme(QLettrage);

    SQL := ' Select E_AUXILIAIRE, E_GENERAL, E_LETTRAGE, E_ETATLETTRAGE, ';
    SQL := SQl + ' E_DEBIT, E_CREDIT, E_DEBITDEV, E_CREDITDEV,E_DATECOMPTABLE, E_DATEPAQUETMIN, ';
    SQL := SQl + ' E_DATEPAQUETMAX, E_COUVERTURE, E_COUVERTUREDEV, E_DEVISE, E_LETTRAGEDEV ';
    SQL := SQl + ' From ECRITURE ';
    SQL := SQl + ' Where ';
    SQL := SQl + ' E_QUALIFPIECE="N" ';
    SQL := SQl + ' And e_eche="X" and e_etatlettrage<>"" and e_etatlettrage<>"RI" ';
    SQL := SQl + ' And e_Ecranouveau<>"OAN" and e_ecranouveau<>"C" ';
    if (VH^.ExoV8.Code <> '') then SQL := SQl + ' And E_DATECOMPTABLE>="' + UsDateTime(VH^.ExoV8.Deb) + '" ';
    if (PourImport) then SQL := SQl + ' And IE_TYPEECR="E" and IE_SELECTED="X" and IE_OKCONTROLE<>"D" '
    else if (VCrit.Etab <> '') then SQL := SQl + ' And E_ETABLISSEMENT="' + VCrit.Etab + '" ';
    { Construction de la clause Order By de la SQL }
    SQL := SQl + ' order BY E_AUXILIAIRE, E_GENERAL, E_ETATLETTRAGE, E_LETTRAGE, E_DEVISE ';

    QLettrage := OpenSQL(SQL,true);
    Rien := QLettrage.Eof;

    Application.ProcessMessages;

{----------------------------------------------------------------------------------------------}

    if (not (QLetDev = nil)) then Ferme(QLetDev);

    SQL := ' Select ' + FE + '_LETTRAGE, ' + FE + '_AUXILIAIRE, ' + FE + '_GENERAL, ' + FE + '_ETATLETTRAGE, ' + FE + '_LETTRAGEDEV ';
    SQL := SQl + ' From ' + FICECR + ' ';
    SQL := SQl + ' Where ';
    SQL := SQl + ' ' + FE + '_QUALIFPIECE="N" and E_LETTRAGE<>"" ';
    SQL := SQl + ' And ' + FE + '_eche="X" and ' + FE + '_etatlettrage<>"" and ' + FE + '_etatlettrage<>"RI" ';
    SQL := SQl + ' And ' + FE + '_Ecranouveau<>"OAN" and ' + FE + '_ecranouveau<>"C" ';
    if (PourImport) then SQL := SQl + ' And IE_TYPEECR="E" and IE_SELECTED="X" and IE_OKCONTROLE<>"D" ';
    if (VCrit.Etab <> '') then SQL := SQl + ' And ' + FE + '_ETABLISSEMENT="' + VCrit.Etab + '" ';
    if (VH^.ExoV8.Code <> '') then SQL := SQl + ' And ' + FE + '_DATECOMPTABLE>="' + UsDateTime(VH^.ExoV8.Deb) + '" ';
    { Construction de la clause Group By de la SQL }
    SQL := SQl + ' Group BY ' + FE + '_AUXILIAIRE, ' + FE + '_GENERAL, ' + FE + '_ETATLETTRAGE, ' + FE + '_LETTRAGE,' + FE + '_LETTRAGEDEV ';

    QLetDev := OpenSQL(SQL,true);
    Rien := Rien and QLetDev.Eof;

    Application.ProcessMessages;

    Result := not Rien;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.RecupCrit;
begin
    Fillchar(VCrit, SizeOf(VCrit), #0);

    with VCrit do
    begin
        Joker := FJalJoker.Visible;
        Etab := '';
        if Joker then
        begin
            Cpt1 := FJalJoker.Text;
            Cpt2 := FJalJoker.Text;
        end
        else
        begin
            Cpt1 := FJal1.Text;
            Cpt2 := FJal2.Text;
        end;

        Date1 := StrToDate(FDateCpta1.Text);
        Date2 := StrToDate(FDateCpta2.Text);
        if FEtab.ItemIndex <> 0 then Etab := FEtab.Value;
        Exo.Code := FExercice.Value;
        if FTypeEcriture.ItemIndex <> 0 then QualifPiece := FTypeEcriture.Value;
        if FNumPiece1.Text <> '' then GL.NumPiece1 := StrToInt(FNumPiece1.Text)
        else GL.NumPiece1 := 0;
        if FNumPiece2.Text <> '' then GL.NumPiece2 := StrToInt(FNumPiece2.Text)
        else GL.NumPiece2 := 999999999;
    end;

    if not DejaLance then
    begin
        Fillchar(VCritLance, SizeOf(VCritLance), #0);
        VCritLance := VCrit;
    end
    else
        LanceDirect := ((VCritLance.Joker = VCrit.Joker) and (VCritLance.Cpt1 = VCrit.Cpt1) and (VCritLance.Cpt2 = VCrit.Cpt2) and
            (VCritLance.Date1 = VCrit.Date1) and (VCritLance.Date2 = VCrit.Date2) and (VCritLance.Etab = VCrit.Etab) and
            (VCritLance.Exo.Code = VCrit.Exo.Code) and (VCritLance.QualifPiece = VCrit.QualifPiece) and
            (VCritLance.GL.NumPiece1 = VCrit.GL.NumPiece1) and (VCritLance.GL.NumPiece2 = VCrit.GL.NumPiece2));
end;

//********** TRAITEMENT PRINCIPAL

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.LanceVerif;
begin
    TNBError1.Caption := '';
    Application.ProcessMessages;

    if (FVerification.ItemIndex = 0) or (FVerification.ItemIndex = 1) then {Verif=Toutes ou Verif Comptable}
    begin
        if not StopVerif then
            if CtrlEcrGen.Checked then SqlEquilEcr(1);
        if not StopVerif then
            if CtrlEcrAna.Checked then SqlEquilEcr(2);
        if not StopVerif then
        begin
            if (not EnBatch) and (not OkVerif) (*and V_PGI.SAV*) then
            begin
                if (MsgRien.Execute(6, '', '') <> mryes) then Exit
                else
                    if CtrlDetGen.Checked then SQLEcr;
            end
            else
                if CtrlDetGen.Checked then SQLEcr;
        end;

        if not StopVerif then
            if CtrlDetAna.Checked then SqlAnal;

        if FVerification.ItemIndex = 0 then {Toutes}
        begin
            LetListe := TStringList.Create;
            if not StopVerif then SQLLet;
            if not StopVerif then SQLVerifLet(1);
            if not StopVerif then SQLVerifLet(2);
        end;
        FiniMove;
    end;

    if not StopVerif and ((FVerification.ItemIndex = 0) or (FVerification.ItemIndex = 3)) then
    begin
        SqlEcrBud;
        FiniMove;
    end;

    if FVerification.ItemIndex = 2 then {Que Verif Lettrage}
    begin
        LetListe := TStringList.Create;
        if not StopVerif then SQLLet;
        if not StopVerif then SQLVerifLet(1);
        if not StopVerif then SQLVerifLet(2);
        FiniMove;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.lanceVerifAnaOff;
var
    NbErrAnaOff : Integer;
begin
    TTravail.Caption := MsgBar.Mess[23];

    Application.ProcessMessages;
{$IFDEF COMPTA}
    NbErrAnaOff := ChercheAnaSansGene(VCrit, LaListe, FALSE);
{$ENDIF}
    if NbErrAnaOff > 0 then
    begin
        OkVerif := FALSE;
        NbError := NbError + NbErrAnaOff;
        TuFaisQuoi(nil, 4);
        Application.ProcessMessages;
    end;

    TTravail.Caption := '';
    Application.ProcessMessages;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.Reparation;
begin
    StopVerif := not MAJ;
    TNBError1.Caption := MsgBar.Mess[20];
    NbError := 0;
    ErrEcr := 0;
    ErrMvtA := 0;
    ErrMvtB := 0;
    ErrAnal := 0;
    ErrMvt := 0;
    ErrLet := 0;

    Application.ProcessMessages;

    OkVerif := True;

    //MAJ:=False ; Exit ;// En Attendant de bien finir la verif....
    TNBError1.Caption := MsgBar.Mess[16];
    //InitMove(NbEnreg,MsgBar.Mess[0]) ;

    if FVerification.ItemIndex <> 2 then {Verif==Toutes ou Verif Comptable}
    begin
        if not StopVerif then SqlEquilEcr(1);
        if not StopVerif then SqlEquilEcr(2);
        if not StopVerif then SQLEcr;
        if not StopVerif then SqlAnal;
        if FVerification.ItemIndex = 0 then
            if not StopVerif then
            begin
                SQLLet;
                SQLVerifLet(1);
                SQLVerifLet(2);
            end; {Toutes}
        //FiniMove ;
    end;

    if FVerification.ItemIndex = 2 then {Que Verif Lettrage}
    begin
        if not StopVerif then
        begin
            SQLLet;
            SQLVerifLet(1);
            SQLVerifLet(2);
        end;
        //FiniMove ;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.TestBreak : Boolean;
begin
    Application.ProcessMessages;
    if StopVerif then
        if MsgRien.Execute(4, '', '') <> mryes then StopVerif := False;
    Result := StopVerif;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.Corrige(Q : TQuery;Champ, Valeur : string);
begin
    Q.Edit;
    Q.FindField(Champ).AsString := Valeur;
    Q.Post;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.InfosCptGen(Ge : string);
var
    i : Integer;
    Inf : TEnregInfos;
begin
    for i := 0 to ListeGen.Count - 1 do
    begin
        if ListeGen.Objects[i] <> nil then
        begin
            Inf := TEnregInfos(ListeGen.Objects[i]);
            if (Inf.Inf1 = Ge) then
            begin
                CGen.General := Inf.Inf1;
                CGen.NatureGene := Inf.Inf2;
                CGen.Collectif := Inf.Bol1;
                CGen.Lettrable := Inf.Bol2;
                CGen.Pointable := Inf.Bol3;
                CGen.Ventilable := Inf.SuperBol;
                Break;
            end;
        end
        else
            Break;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.InfosCptAux(Au : string);
var
    i : Integer;
    Inf : TEnregInfos;
    CAux : TGTiers;
begin
    Caux := nil;
    for i := 0 to ListeAux.Count - 1 do
    begin
        if ListeAux.Objects[i] <> nil then
        begin
            Inf := TEnregInfos(ListeAux.Objects[i]);
            if (Inf.Inf1 = Au) then
            begin
                CTiers.Auxi := Inf.Inf1;
                CTiers.NatureAux := Inf.Inf2;
                CTiers.Lettrable := Inf.Bol1;
                Break;
                Exit;
            end;
        end
        else
        begin
            Break;
            exit;
        end;
    end;
    if (ListeAux.Count) > 2000 then
    begin
        if CAux = nil then CAux := TGTiers.Create(Au);
        if CAux <> nil then
        begin
            CTiers.Auxi := CAux.Auxi;
            CTiers.NatureAux := CAux.NatureAux;
            CTiers.Lettrable := CAux.Lettrable;
        end;
        if CAux <> nil then
        begin
            CAux.FREE;
        end;
    end
    else
        Exit;
end;

//********** Etape du TRAITEMENT

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.SQLEcr;
var
    SiLet, SiPoint, SiVent, MEcr : Boolean;
    IdentPiece : tIdentPiece;
    PremFois : Boolean;
    PbDate : Boolean;
begin
    SiLet := False;
    SiPoint := False;
    SiVent := False;
    MEcr := True;
    QEcr.First;
    PremFois := TRUE;

    Fillchar(IdentPiece, SizeOf(IdentPiece), #0);

    while not QEcr.Eof do
    begin
        if Mvt <> nil then
        begin
            Mvt.FREE;
            Mvt := nil;
        end;

        MoveCur(False);
        EnregMvtEcr;

        if PremFois then
        begin
            MvtToIdent(IdentPiece, Mvt);
            {JP 17/10/03 : ça ne coûte pas très cher de renseigner la nature de pièce et ça permet de s'assurer
                           que tous les mouvements d'une même pièce ont la même nature de pièce}
            IdentPiece.NatP := Mvt.NATUREPIECE;
        end;

        if RupturePiece(IdentPiece, Mvt) then
        begin
            MvtToIdent(IdentPiece, Mvt);
            {JP 17/10/03 : ça ne coûte pas très cher de renseigner la nature de pièce et ça permet de s'assurer
                           que tous les mouvements d'une même pièce ont la même nature de pièce}
            IdentPiece.NatP := Mvt.NATUREPIECE;
        end
        else
        begin
            PbDate := (IdentPiece.DateP <> Mvt.DateComptable) and (Mvt.ModeSaisie <> 'BOR') and (Mvt.ModeSaisie <> 'LIB') and PourImport;
            if PbDate then
            begin
                LaListeAdd(GoListe1(Mvt, 'C', MsgLibel.Mess[112] + ' ' + IntToStr(Mvt.NUMEROPIECE) + ' ' + MsgLibel.Mess[113], 255, 'G'));
                OkVerif := FALSE;
            end;
        end;

        PremFois := FALSE;

        if VerifCompte(SiLet, SiPoint, SiVent) then
        begin
            if VerifInfoJal then
            begin
                if VerifInfoPerio then
                begin
                    if VerifInfoExo then
                    begin
                        if VerifInfoEcrAno then
                        begin
                            if Mvt.ECRANOUVEAU <> 'H' then
                            begin
                                if not VerifInfoEche(SiLet, SiPoint) then OkVerif := False;
                                if not VerifInfoDevise then OkVerif := False;
                                if not VerifInfoPieces then OkVerif := False;

                                {JP 17/10/03 : On s'assure que le mouvement a la même nature que que la pièce }
                                {CA 19/12/03 : en mode pièce uniquement }
                                if ((Mvt.MODESAISIE = '-') and (IdentPiece.NatP <> Mvt.NATUREPIECE)) then
                                begin
                                    LaListeAdd(GoListe1(Mvt, 'C', ' (' + Mvt.NATUREPIECE + ' au lieu de ' + IdentPiece.NatP + ') ', 115, 'E'));
                                    OkVerif := False;
                                end;

                                if not VerifTauxDev(MEcr) then OkVerif := False;
                            end;
                            if not VerifInfoAnal(SiVent) then OkVerif := False;
                            //if not VerifInfoConf then OkVerif:=False ;
                            if not VerifEtab then OkVerif := False;
                        end
                        else OkVerif := False;
                    end
                    else OkVerif := False;
                end
                else OkVerif := False;
            end
            else OkVerif := False;
        end
        else OkVerif := False;

        TuFaisQuoi(QEcr, 3);
        if TestBreak then Break;
        QEcr.Next;
    end;
    //StopC(5) ;ShowC(5,'Temp ecriture ', False) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.SqlEquilEcr(Ec : Byte); {EC=1 : Verif équilibre Ecriture ; EC=2 : Verif équilibre Ecr-Ana}
begin
    if Ec = 1 then
    begin
        { Somme D et C (Devise pivot et devise) de chaque Pièce }
        QSum.First;
        Fillchar(INFOSPIECE, SizeOf(INFOSPIECE), #0);

        while not QSum.Eof do
        begin
            MoveCur(False);
            if Mvt <> nil then
            begin
                Mvt.FREE;
                Mvt := nil;
            end;
            EnregEquil(QSum, Ec);
            TuFaisQuoi(QSum, 1);
            if not EquilEcr then OkVerif := False; { Test si la pièce est équilibrée : ECRITURE seule }
            if TestBreak then Break;
            QSum.Next;
        end;
    end
    else
    begin
        { Somme D et C (Devise pivot et devise) de chaque Pièce/Ligne pour vérif ECR-ANA}
        QSumGeAn.First;
        // suppresion des deux fonction prepare .....
//        PrepareAnal;
//        PrepareAutres;

        while not QSumGeAn.Eof do
        begin
            MoveCur(False);
            if Mvt <> nil then
            begin
                Mvt.FREE;
                Mvt := nil;
            end;
            EnregEquilA;
            TuFaisQuoi(QSumGeAn, 2);
            if not EquilAnal then OkVerif := False; {Test de l' équilibre la pièce/Ligne ECRITURE Avec la pièce/Ligne de L'ANALYTIQUE }
            if TestBreak then Break;
            QSumGeAn.Next;
        end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.EnregMvtEcr;
var
    MS : string;
begin {Pour Verif aussi avant imports}
    Mvt := TInfoMvt.Create;
    Mvt.JOURNAL := QEcr.Fields[0].AsString;
    Mvt.DATECOMPTABLE := QEcr.Fields[1].AsDateTime;
    Mvt.NUMLIGNE := QEcr.Fields[2].AsInteger;
    Mvt.GENERAL := QEcr.Fields[3].AsString;
    Mvt.AUXILIAIRE := QEcr.Fields[4].AsString;
    Mvt.DEVISE := QEcr.Fields[5].AsString;
    Mvt.MODEPAIE := QEcr.Fields[6].AsString;
    Mvt.DEBIT := QEcr.Fields[7].AsFloat;
    Mvt.CREDIT := QEcr.Fields[8].AsFloat;
    Mvt.DEBITDEV := QEcr.Fields[9].AsFloat;
    Mvt.CREDITDEV := QEcr.Fields[10].AsFloat;
    Mvt.ANA := QEcr.Fields[11].AsString;
    Mvt.REFINTERNE := QEcr.Fields[12].AsString;
    Mvt.DATEECHEANCE := QEcr.Fields[13].AsDateTime;
    Mvt.NUMECHE := QEcr.Fields[14].AsInteger;
    Mvt.ETABLISSEMENT := QEcr.Fields[15].AsString;
    Mvt.ETATLETTRAGE := QEcr.Fields[16].AsString;
    Mvt.ECHE := QEcr.Fields[17].AsString;
    Mvt.ECRANOUVEAU := QEcr.Fields[18].AsString;
    Mvt.COUVERTURE := QEcr.Fields[19].AsFloat;
    Mvt.COUVERTUREDEV := QEcr.Fields[20].AsFloat;
    Mvt.NATUREPIECE := QEcr.Fields[21].AsString;
    if not PourImport then
    begin
        Mvt.NUMEROPIECE := QEcr.Fields[22].AsInteger;
        Mvt.CONFIDENTIEL := QEcr.Fields[23].AsString;
        Mvt.EXERCICE := QEcr.Fields[24].AsString;
        Mvt.MODESAISIE := QEcr.Findfield('' + FE + '_MODESAISIE').AsString;
    end
    else
    begin
        Mvt.NUMEROPIECE := QEcr.Fields[22].AsInteger;
        Mvt.CHRONO := QEcr.Fields[23].AsInteger;
        Mvt.QUALIFPIECE := QEcr.Fields[25].AsString;
        Mvt.EXERCICE := 'IMP';
        MS := QEcr.Findfield('' + FE + '_LIBREBOOL1').AsString;
        if MS <> '' then
            case MS[1] of
                'L' : Mvt.MODESAISIE := 'LIB';
                'B' : Mvt.MODESAISIE := 'BOR';
                else Mvt.MODESAISIE := '-';
            end;
    end;
    Mvt.TAUXDEV := QEcr.Findfield('' + FE + '_TAUXDEV').AsFloat;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.EnregEquil(Q : TQuery;Quelle : byte); { Enregistrement des valeurs de la requete...}
var
    DecalImport : Byte;
    St : string;
    D, M, Y : Word;
begin { ...Pour la vérif de l'équilibre sur la table ECRITURE seul !!!}
    Mvt := TInfoMvt.Create;

    Mvt.JOURNAL := Q.Fields[0].AsString; // journal
    DecalImport := 1;

    if PourImport then
    begin
        DecalImport := 0;
        St := Q.Fields[1 + DecalImport].AsString;
        D := 1;
        M := StrToInt(Copy(St, 3, 2));
        Y := StrToInt(Copy(St, 1, 2));
        if Y > 80 then Y := 1900 + Y
        else Y := 2000 + Y;
        Mvt.DATECOMPTABLE := EncodeDate(Y, M, D);
    end
    else Mvt.DATECOMPTABLE := Q.Fields[1 + DecalImport].AsDateTime;

    Mvt.DEVISE := Q.Fields[3 + DecalImport].AsString;
    Mvt.DEBIT := Q.Fields[4 + DecalImport].AsFloat;
    Mvt.CREDIT := Q.Fields[5 + DecalImport].AsFloat;
    Mvt.DEBITDEV := Q.Fields[6 + DecalImport].AsFloat;
    Mvt.CREDITDEV := Q.Fields[7 + DecalImport].AsFloat;

    if not PourImport then
    begin
        Mvt.NUMEROPIECE := Q.Fields[3].AsInteger;
        Mvt.EXERCICE := Q.FieldS[1].AsString;
    end
    else
    begin
        Mvt.NUMEROPIECE := Q.FieldS[2].AsInteger;
        Mvt.EXERCICE := 'IMP';
    end;

    Mvt.DEBITMin := Q.Fields[8 + DecalImport].AsFloat;
    Mvt.DEBITMax := Q.Fields[9 + DecalImport].AsFloat;
    Mvt.CREDITMin := Q.Fields[10 + DecalImport].AsFloat;
    Mvt.CREDITMax := Q.Fields[11 + DecalImport].AsFloat;

    if PourImport
      then Mvt.MAXCONTROLE := Q.Fields[12].AsString
      else Mvt.NATUREPIECE := Q.Fields[13].AsString ; // FQ17137 : Chargement de la Nature de Pièce

end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.EquilEcr : Boolean; { Test de l'équilibre sur la Table ECRITURE seul !!!}
var
    CBonP, CBonD, CBonE, Resultat, OkOK, CBonPAbs : Boolean;
    MD, MC, MDD, MCD, MDE, MCE : Double;
    DecDev : Byte;
//    NoEcart          : Boolean ;
begin
(*
    NoEcart := FALSE;

    if Not (ctxPCL in V_PGI.PGIContexte) then
    begin
        if EstSpecif('51188') then
        begin
            if StopEcart(IdentPiece.DateP) Then NoEcart := TRUE;
        end;
    end;
*)

    Resultat := True;
    DecDev := 0;
    OkOK := True;

    if ((INFOSPIECE.Jal <> Mvt.Journal) or (INFOSPIECE.Exo <> Mvt.Exercice) and (INFOSPIECE.Dat <> Mvt.DateComptable) or (INFOSPIECE.NumP <> Mvt.NumeroPiece)) then
    begin
        if Mvt.DEVISE = '' then OkOk := False;
    end
    else OkOk := False;

    INFOSPIECE.Jal := Mvt.Journal;
    INFOSPIECE.Exo := Mvt.Exercice;
    INFOSPIECE.Dat := Mvt.DateComptable;
    INFOSPIECE.NumP := Mvt.NumeroPiece;

    if not OkOk then
    begin
        Result := Resultat;
        Exit;
    end;

    { Devise Pivot : Pièce équilibrée ? }
    MD := Arrondi(Mvt.DEBIT, V_PGI.OkdecV);
    MC := Arrondi(Mvt.CREDIT, V_PGI.OkdecV);
    CBonP := (MD = MC);

    { Devise pivot : Pièce à 0 ? }
    CBonPAbs := True;
    if (Arrondi(Mvt.DEBITMIN, V_PGI.OkdecV) = 0) and (Arrondi(Mvt.DEBITMAX, V_PGI.OkdecV) = 0) and (Arrondi(Mvt.CREDITMIN, V_PGI.OkdecV) = 0) and (Arrondi(Mvt.CREDITMAX, V_PGI.OkdecV) = 0) then CBonPAbs := False;

    { Devise : Pièce équilibrée ? }
    if Mvt.DEVISE <> V_PGI.DevisePivot then
    begin
        TrouveDecDev(Mvt.DEVISE, DecDev);
        MDD := Arrondi(Mvt.DEBITDEV, DecDev);
        MCD := Arrondi(Mvt.CREDITDEV, DecDev);
        // CriCri : 02/07/97
        // DecDev -> DecEuro
    end
    else
    begin
        MDD := Arrondi(Mvt.DEBITDEV, V_PGI.OkdecV);
        MCD := Arrondi(Mvt.CREDITDEV, V_PGI.OkdecV);
        // CriCri : 02/07/97
        // DecDev -> DecEuro
       //CBonD:=(MDD=MCD) ;
    end;
    CBonD := (MDD = MCD);
    CBonE := False;
    MDE := 0;
    MCE := 0;

    { Pièce à zéro en Devise Pivot ; Affichage dans la grille }
    if not CBonPAbs then // A 1?
    begin
        if not MAJ then ErrEquilEcr(Resultat, 6, MD, MC, MDD, MCD, MDE, MCE);
    end
    { Pièce non équilibré en Devise Pivot et en Devise ; Affichage dans la grille }
    else
        if (not CBonP) and (not CBonD) then // A 13
    begin
        if not MAJ then ErrEquilEcr(Resultat, 1, MD, MC, MDD, MCD, MDE, MCE);
    end
    { Pièce non équilibré en Devise Pivot ; Affichage dans la grille }
    else
        if not CBonP then // A 11
    begin
        if not MAJ then ErrEquilEcr(Resultat, 2, MD, MC, MDD, MCD, MDE, MCE);
    end
    { Pièce non équilibré en Devise ; Affichage dans la grille }
    else
        if not CBonD then // A 12
    begin
        if not MAJ then ErrEquilEcr(Resultat, 3, MD, MC, MDD, MCD, MDE, MCE);
    end;

    { Si la devise est renseigné les montants en DEBITDEV et CREDITDEV doivent etre <> 0 }
    if Mvt.DEVISE <> V_PGI.DevisePivot then
      if (Mvt.NATUREPIECE <> 'ECC') then // FQ
        if ((Mvt.DEBITDEV = 0) and (Mvt.CREDITDEV = 0)) then
        begin
            if not MAJ then ErrEquilEcr(Resultat, 4, MD, MC, 0, 0, 0, 0);
        end;

    if PourImport and CBonP and CBonD and CBonE and (Mvt.MAXCONTROLE = 'X') then
    begin
        if not MAJ then ErrEquilEcr(Resultat, 7, MD, MC, MDD, MCD, MDE, MCE);
    end;

    Result := Resultat;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.ErrEquilEcr(var Err : Boolean;E : Byte;D, C, DD, CD, DE, CE : Double);
var
    Q : TQuery; {Recherche RefInterne sur Piece Non Equil}
    T : TQUERY; {Corrections si possible }
    st : string;
    MontantPivot, MontantDevise {, MontantEuro} : Double;
begin
    T := nil;

    St := ' Select ' + FE + '_REFINTERNE, ' + FE + '_NUMLIGNE, ' + FE + '_NATUREPIECE, ' + FE + '_QUALIFPIECE, ' + FE + '_ETABLISSEMENT, ';
    St := St + ' ' + FE + '_DEBITDEV, ' + FE + '_CREDITDEV, ' + FE + '_DEVISE, ' + FE + '_DATECOMPTABLE ';
    if PourImport then st := st + ', IE_CHRONO, IE_NUMPIECE '
    else st := st + ', E_EXERCICE, E_NUMEROPIECE ';
    St := St + ' From ' + FICECR + ' where ' + FE + '_JOURNAL="' + Mvt.Journal + '" ';
    if not PourImport then St := St + ' AND ' + FE + '_EXERCICE="' + Mvt.Exercice + '" ';
    St := St + ' AND ' + FE + '_DATECOMPTABLE="' + USDATETIME(Mvt.DATECOMPTABLE) + '" ';
    if not PourImport then St := St + ' AND ' + FE + '_NUMEROPIECE=' + IntToStr(Mvt.NUMEROPIECE) + ' '
    else St := St + ' AND ' + FE + '_NUMPIECE=' + IntToStr(Mvt.NUMEROPIECE) + ' AND ' + FE + '_DEVISE="' + Mvt.DEVISE + '" ';
    if not PourImport then St := St + ' Order by E_JOURNAL, E_EXERCICE,  E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE , E_DEVISE '
    else St := St + ' Order by IE_JOURNAL, IE_DATECOMPTABLE, IE_NUMPIECE, IE_NUMLIGNE , IE_DEVISE ';

    Q := OpenSql(St, True);

    Mvt.REFINTERNE := Q.FindField('' + FE + '_REFINTERNE').AsString;
    Mvt.NUMLIGNE := Q.FindField('' + FE + '_NUMLIGNE').AsInteger;
    if PourImport then Mvt.CHRONO := Q.FindField('IE_CHRONO').AsInteger;

    if MAJ then
    begin
        T := OpenSQL('SELECT * FROM ECRITURE WHERE E_GENERAL="Ed#"', False);
        T.Insert;
        InitNew(T);
        T.FindField('E_JOURNAL').AsString := Mvt.Journal;
        T.FindField('E_EXERCICE').AsString := Mvt.Exercice;
        T.FindField('E_DATECOMPTABLE').AsDateTime := Mvt.DATECOMPTABLE;
        T.FindField('E_NUMEROPIECE').AsInteger := Mvt.NUMEROPIECE;
        T.FindField('E_NUMLIGNE').AsInteger := Q.FindField('E_NUMLIGNE').AsInteger + 1;
        T.FindField('E_DEVISE').AsString := Mvt.DEVISE;
        T.FindField('E_REFINTERNE').AsString := Mvt.REFINTERNE;
        T.FindField('E_GENERAL').AsString := VH^.Cpta[fbGene].Attente;
        T.FindField('E_NATUREPIECE').AsString := Q.FindField('E_NATUREPIECE').AsSTring;
        T.FindField('E_QUALIFPIECE').AsString := Q.FindField('E_QUALIFPIECE').AsSTring;
        T.FindField('E_ETABLISSEMENT').AsString := Q.FindField('E_ETABLISSEMENT').AsSTring;
        T.FindField('E_LIBELLE').AsString := MsgLibel2.Mess[2];
    end;

    Err := False;

    case E of {E --> 1 : Pivot + Dev ; 2 : Pivot ; 3 : Devise ; 4 : Piece Devise avec Monants devise=0 ; 5 : Euro ; 6 : Pièce à 0}
        1 :
            if MAJ and (T <> nil) then
            begin
                MontantPivot := LeSolde(D, C);
                MontantDevise := LeSolde(DD, CD); //MontantEuro:=LeSolde(DE,CE) ;

                if MontantDevise < 0 then
                begin
                    T.FindField('E_DEBITDEV').AsFloat := 0;
                    T.FindField('E_CREDITDEV').AsFloat := Abs(MontantDevise);
                end
                else
                begin
                    T.FindField('E_DEBITDEV').AsFloat := Abs(MontantDevise);
                    T.FindField('E_CREDITDEV').AsFloat := 0;
                end;

                if MontantPivot < 0 then
                begin
                    T.FindField('E_DEBIT').AsFloat := 0;
                    T.FindField('E_CREDIT').AsFloat := Abs(MontantPivot);
                end
                else
                begin
                    T.FindField('E_DEBIT').AsFloat := Abs(MontantPivot);
                    T.FindField('E_CREDIT').AsFloat := 0;
                end;
            end
            else
            begin
                LaListeAdd(GoListe1(Mvt, 'C', MsgLibel.Mess[15] + FloatToStr(D) + '/' + FloatToStr(DD) + ' ; ' + MsgLibel.Mess[16] + FloatToStr(C) + '/' + FloatToStr(CD), 2, 'G'));
            end;
        2 :
            if MAJ then
            begin
                MontantPivot := LeSolde(D, C);
                if MontantPivot < 0 then
                begin
                    T.FindField('E_DEBIT').AsFloat := 0;
                    T.FindField('E_CREDIT').AsFloat := Abs(MontantPivot);
                end
                else
                begin
                    T.FindField('E_DEBIT').AsFloat := Abs(MontantPivot);
                    T.FindField('E_CREDIT').AsFloat := 0;
                end;
            end
            else
            begin
                LaListeAdd(GoListe1(Mvt, 'C', MsgLibel.Mess[15] + FloatToStr(D) + ' ; ' + MsgLibel.Mess[16] + FloatToStr(C), 0, 'G'));
            end;
        3 :
            if MAJ then
            begin
                MontantDevise := LeSolde(DD, CD);
                if MontantDevise < 0 then
                begin
                    T.FindField('E_DEBITDEV').AsFloat := 0;
                    T.FindField('E_CREDITDEV').AsFloat := Abs(MontantDevise);
                end
                else
                begin
                    T.FindField('E_DEBITDEV').AsFloat := Abs(MontantDevise);
                    T.FindField('E_CREDITDEV').AsFloat := 0;
                end;
            end
            else
            begin
                LaListeAdd(GoListe1(Mvt, 'C', MsgLibel.Mess[15] + FloatToStr(DD) + ' ; ' + MsgLibel.Mess[16] + FloatToStr(CD), 1, 'G'));
            end;

        7 :
            if MAJ then
            begin
                MontantPivot := LeSolde(D, C);
                if MontantPivot < 0 then
                begin
                    T.FindField('E_DEBIT').AsFloat := 0;
                    T.FindField('E_CREDIT').AsFloat := Abs(MontantPivot);
                end
                else
                begin
                    T.FindField('E_DEBIT').AsFloat := Abs(MontantPivot);
                    T.FindField('E_CREDIT').AsFloat := 0;
                end;
            end
            else
            begin
                LaListeAdd(GoListe1(Mvt, 'C', MsgLibel.Mess[106] + ' ' + IntToStr(Mvt.NUMEROPIECE) + ' ' + MsgLibel.Mess[107], 255, 'G'));
            end;
    end;

    if MAJ then
    begin
        T.Post;
        Ferme(T);
    end;

    Ferme(Q);
end;

//********** Etape du TRAITEMENT Analytique

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.SqlAnal;
var
    MEcr : Boolean;
begin
    { Verif des Mvts Analytique }
    QAnal.First;
    MEcr := False;

    while not QAnal.Eof do
    begin
        if Mvt <> nil then
        begin
            Mvt.FREE;
            Mvt := nil;
        end;

        MoveCur(False);
        EnregMvtAna;

        if VerifCptAnal then
        begin
            if VerifInfoJalAnal then
            begin
                if VerifInfoExoAnal then
                begin
                    if VerifInfoEcrAnoAnal then
                    begin
                        if Mvt.ECRANOUVEAU <> 'H' then
                        begin
                            if not VerifInfoPieceAnal then OkVerif := False;
                            //if not VerifInfoConfAnal then OkVerif:=False ;
                            if not VerifInfoDeviseAnal then OkVerif := False;
                            if not VerifEtabAnal then OkVerif := False;
                            if not VerifTauxDev(MEcr) then OkVerif := False;
                        end;
                    end
                    else OkVerif := False;
                end
                else OkVerif := False;
            end
            else OkVerif := False;
        end
        else OkVerif := False;

        TuFaisQuoi(QAnal, 4);
        if TestBreak then Break;
        //If MAJ then Corrige ;
        QAnal.Next;
    end;
    //StopC(4) ;ShowC(4,'Temp Analytique', False) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.EnregMvtAna; { A cause de PourVérif (Compt. ou Impor) CAD Prefixe de table}
begin
    Mvt := TInfoMvt.Create;
    Mvt.JOURNAL := QAnal.Fields[0].AsString;
    Mvt.DATECOMPTABLE := QAnal.Fields[1].AsDateTime;
    Mvt.NUMLIGNE := QAnal.Fields[2].AsInteger;
    Mvt.Axe := QAnal.Fields[3].AsString;
    Mvt.DEVISE := QAnal.Fields[4].AsString;
    Mvt.REFINTERNE := QAnal.Fields[5].AsString;
    Mvt.GENERAL := QAnal.Fields[6].AsString;
    Mvt.NUMVENTIL := QAnal.Fields[7].AsInteger;
    Mvt.SECTION := QAnal.Fields[8].AsString;
    Mvt.ECRANOUVEAU := QAnal.Fields[9].AsString;
    Mvt.NATUREPIECE := QAnal.Fields[10].AsString;
    Mvt.ETABLISSEMENT := QAnal.Fields[11].AsString;
    Mvt.TYPEANALYTIQUE := QAnal.Fields[12].AsString;
    Mvt.DEBIT := QAnal.Fields[13].AsFloat;
    Mvt.CREDIT := QAnal.Fields[14].AsFloat;
    Mvt.DEBITDEV := QAnal.Fields[15].AsFloat;
    Mvt.CREDITDEV := QAnal.Fields[16].AsFloat;

    if not PourImport then
    begin
        Mvt.NUMEROPIECE := QAnal.Fields[17].AsInteger;
        Mvt.CONFIDENTIEL := QAnal.Fields[18].AsString;
        Mvt.EXERCICE := QAnal.Fields[19].AsString;
    end
    else
    begin
        Mvt.NUMEROPIECE := QAnal.Fields[17].AsInteger;
        Mvt.CHRONO := QAnal.Fields[18].AsInteger;
        Mvt.EXERCICE := 'IMP';
    end;

    Mvt.TAUXDEV := QAnal.Findfield('' + FY + '_TAUXDEV').AsFloat;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.EquilAnal : Boolean;
var
    CBonP, CBonD : Boolean;
    InfAna : TINFOSMVTANA;
    Tot : TabTot;
    DecDev, I : Byte;
    AxeOk : array[1..5] of Boolean;
    AxeEnCours : string;
    IAxeEnCours : Byte;
    SetAxe : tSetAxe;
    ClePaf : string;
    SQL : string;
    //SG6
    lIntI : integer;
    lVarField : variant;
begin
    Fillchar(InfAna, SizeOf(InfAna), #0);
    Result := True; // B 10
    for i := 1 to 5 do AxeOk[i] := FALSE;
    // Rony 15/05/97
    if Mvt.DEVISE = '' then Exit;

    { Recherche les Piéce/Lignes de type Analytique }
    { Somme des D & C des Pièces Analytique attachés }

//    if PourImport then I := 1
//    else I := 0;
//
//    QAnaGene.Params[0].AsString := Mvt.JOURNAL;
//    if not PourImport then QAnaGene.Params[1].AsString := Mvt.EXERCICE;
//    QAnaGene.Params[2 - I].AsDateTime := Mvt.DATECOMPTABLE;
//    QAnaGene.Params[3 - i].AsInteger := Mvt.NUMEROPIECE;
//    QAnaGene.Params[4 - i].AsInteger := Mvt.NUMLIGNE;
//    QAnaGene.Params[5 - i].AsString := Mvt.DEVISE;
//    QAnaGene.Params[6 - i].AsString := Mvt.Etablissement;
////    QAnaGene.Params[7-i].AsString:=Mvt.AXE ;
//    SetAxe := [];
//    QAnaGene.Params[7 - i].AsString := Mvt.AXEPLUS[1];
//    RecupAxe(Mvt.AXEPLUS[1], SetAxe);
//    QAnaGene.Params[8 - i].AsString := Mvt.AXEPLUS[2];
//    RecupAxe(Mvt.AXEPLUS[2], SetAxe);
//    QAnaGene.Params[9 - i].AsString := Mvt.AXEPLUS[3];
//    RecupAxe(Mvt.AXEPLUS[3], SetAxe);
//    QAnaGene.Params[10 - i].AsString := Mvt.AXEPLUS[4];
//    RecupAxe(Mvt.AXEPLUS[4], SetAxe);
//    QAnaGene.Params[11 - i].AsString := Mvt.AXEPLUS[5];
//    RecupAxe(Mvt.AXEPLUS[5], SetAxe);
//    QAnaGene.Params[12 - i].AsString := Mvt.QUALIFPIECE;

    SetAxe := [];
    RecupAxe(Mvt.AXEPLUS[1], SetAxe);
    RecupAxe(Mvt.AXEPLUS[2], SetAxe);
    RecupAxe(Mvt.AXEPLUS[3], SetAxe);
    RecupAxe(Mvt.AXEPLUS[4], SetAxe);
    RecupAxe(Mvt.AXEPLUS[5], SetAxe);

    if (not (QAnaGene = nil)) then ferme(QAnaGene);

    SQL := ' Select';
    SQL := SQL + ' Sum(' + FY + '_DEBIT) as SumD, Sum(' + FY + '_CREDIT) as SumC, ';
    SQL := SQL + ' Sum(' + FY + '_DEBITDEV) as SumDDEV, Sum(' + FY + '_CREDITDEV) as SumCDEV, ';
    SQL := SQL + ' ' + FY + '_AXE AS THEAXE';
    //SG6 09.03.05    mode croisaxe
    if VH^.AnaCroisaxe and (not PourImport) then
    begin
      SQL := SQL + ', Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5';
    end;
    SQL := SQL + ' From ' + FICANA + ' ';
    SQL := SQL + ' Where ';
    SQL := SQL + ' ' + FY + '_JOURNAL="' + Mvt.JOURNAL + '" ';
    if not PourImport then SQL := SQL + ' And Y_EXERCICE="' + Mvt.EXERCICE + '" ';
    SQL := SQL + ' And ' + FY + '_DATECOMPTABLE="' + UsDateTime(Mvt.DATECOMPTABLE) + '" ';
    if not PourImport then SQL := SQL + ' And Y_NUMEROPIECE=' + IntToStr(Mvt.NUMEROPIECE) + ' '
    else SQL := SQL + ' And IE_NUMPIECE=' + IntToStr(Mvt.NUMEROPIECE) + ' ';
    SQL := SQL + ' And ' + FY + '_NUMLIGNE=' + IntToStr(Mvt.NUMLIGNE) + ' ';
    SQL := SQL + ' And ' + FY + '_DEVISE="' + Mvt.DEVISE + '" ';
    SQL := SQL + ' And ' + FY + '_ETABLISSEMENT="' + Mvt.Etablissement + '" ';
    SQL := SQL + ' And ' + FY + '_AXE IN ("' + Mvt.AXEPLUS[1] + '","' + Mvt.AXEPLUS[2] + '","' + Mvt.AXEPLUS[3] + '","' + Mvt.AXEPLUS[4] + '","' + Mvt.AXEPLUS[5] + '") ';
    SQL := SQL + ' And ' + FY + '_QUALIFPIECE="' + Mvt.QUALIFPIECE + '" ';
    SQL := SQL + ' And ' + FY + '_QUALIFPIECE<>"C" ';
    if PourImport then SQL := SQL + ' And IE_TYPEECR="A" and IE_SELECTED="X" and IE_OKCONTROLE<>"D" ';
    { Construction de la clause Group By de la SQL }
    if not PourImport then
    begin
      if not VH^.AnaCroisaxe then SQL := SQL + ' Group By Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE , Y_AXE, Y_DEVISE '
      else SQL := SQL + ' Group By Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE , Y_AXE, Y_DEVISE, Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5';
    end
    else SQL := SQL + ' Group By IE_JOURNAL, IE_DATECOMPTABLE, IE_NUMPIECE, IE_NUMLIGNE, IE_AXE, IE_DEVISE';

    QAnaGene := OpenSQL(SQL,true);

{ Pour Info ! }
{ ToT[1] --> Mvt Gene , Tot[3] --> Anal En Pivot , Tot}
{ ToT[2] --> Mvt Gene , Tot[4] --> Anal En Devise }
{ ToT[5] --> Mvt Gene , Tot[6] --> Anal En Euro }

    while not QAnaGene.Eof do { L'écriture a t-elle des analytiques rattachées ? }
    begin
        Fillchar(Tot, SizeOf(Tot), #0);
        AxeEnCours := QAnaGene.FindField('THEAXE').AsString;
        //SG6 09.03.05 Gestion mode croisaxe
        if not VH^.AnaCroisaxe then
        begin
          if Length(AxeEnCours) >= 2 then IAxeEnCours := StrToInt(Copy(AxeEnCours, 2, 1))
          else IAxeEnCours := 0;
          if IAxeEnCours in [1..5] then AxeOk[IAxeEnCours] := TRUE;
        end
        else
        begin
          for lIntI := 1 to 5 do
          begin
            lVarField := QAnaGene.FindField('Y_SOUSPLAN' + IntToStr(lIntI)).AsVariant;
            if VarIsNull(lVarField) then continue;
            if QAnaGene.FindField('Y_SOUSPLAN' + IntToStr(lIntI)).AsString <> '' then AxeOk[lIntI] := True;
          end;
        end;


        //if PourImport then Mvt.CHRONO:=QAnaGene.FindField('IE_CHRONO').AsInteger ;
        { Les Montants des Pièces/Lignes Ecr == aux montants Pièces/Lignes analytque ? ... }
        { ...En Devise Pivot }
        Tot[1].TotDebit := Arrondi(Mvt.DEBIT, V_PGI.OkdecV);
        Tot[1].TotCredit := Arrondi(Mvt.CREDIT, V_PGI.OkdecV);
        Tot[3].TotDebit := Arrondi(QAnaGene.FindField('SumD').AsFloat, V_PGI.OkdecV);
        Tot[3].TotCredit := Arrondi(QAnaGene.FindField('SumC').AsFloat, V_PGI.OkdecV);
        { ...En Devise }

        if Mvt.DEVISE <> V_PGI.DevisePivot then
        begin
            DecDev := 0;
            TrouveDecDev(Mvt.DEVISE, DecDev);
            Tot[2].TotDebit := Arrondi(Mvt.DEBITDEV, DecDev);
            Tot[2].TotCredit := Arrondi(Mvt.CREDITDEV, DecDev);
            Tot[4].TotDebit := Arrondi(QAnaGene.FindField('SumDDev').AsFloat, DecDev);
            Tot[4].TotCredit := Arrondi(QAnaGene.FindField('SumCDev').AsFloat, DecDev);
        end
        else
        begin
            Tot[2].TotDebit := Arrondi(Mvt.DEBITDEV, V_PGI.OkdecV);
            Tot[2].TotCredit := Arrondi(Mvt.CREDITDEV, V_PGI.OkdecV);
            Tot[4].TotDebit := Arrondi(QAnaGene.FindField('SumDDev').AsFloat, V_PGI.OkdecV);
            Tot[4].TotCredit := Arrondi(QAnaGene.FindField('SumCDev').AsFloat, V_PGI.OkdecV);
        end;

        Tot[5].TotDebit := 0;
        Tot[5].TotCredit := 0;
        Tot[6].TotDebit := 0;
        Tot[6].TotCredit := 0;

        CBonP := (Tot[1].TotDebit = Tot[3].TotDebit) and (Tot[1].totCredit = Tot[3].totCredit);
        CBonD := (Tot[2].TotDebit = Tot[4].TotDebit) and (Tot[2].totCredit = Tot[4].TotCredit);

        //If V_PGI.SAV then CBonE:=(Tot[5].TotDebit=Tot[6].TotDebit) and (Tot[5].totCredit=Tot[6].TotCredit)
        //Else CBonE:=True ;
        { Dev. Pivot et Dev. : affichage dans la liste des Montants des Pièces/Lignes Ecr <> aux montants Pièces/Lignes analytque ? ... }
        if (not CBonP) and (not CBonD) then // B 14
        begin
            if ShuntPbAna and PourImport and PourNewImport and (ListePbAna <> nil) then
            begin
                PbAna := TRUE;
                ClePAF := ClePieceAnaFausse(MVT);
                ListePbAna.Add(ClePAF);
            end
            else
            begin
                if not MAJ then
                begin
                    LaListeAdd(GoListe1(Mvt, 'C', AxeEnCours + ' : ' + MsgLibel.Mess[15] + FloatToStr(Tot[3].TotDebit) + ' (' + FloatToStr(Tot[1].TotDebit) + ') '
                        + MsgLibel.Mess[16] + FloatToStr(Tot[3].TotCredit) + ' (' + FloatToStr(Tot[1].TotCredit) + ') '
                        + '& ' + MsgLibel.Mess[15] + FloatToStr(Tot[4].TotDebit) + ' (' + FloatToStr(Tot[2].TotDebit) + ') '
                        + MsgLibel.Mess[16] + FloatToStr(Tot[4].TotCredit) + ' (' + FloatToStr(Tot[2].TotCredit) + ') '
                        , 5, 'A'));
                    Result := False;
                end
                else CorrEquilAna(1, Tot);
            end;
        end
        else
            if not CBonP then // B 12 { Dev. Pivot : affichage dans la liste des Montants des Pièces/Lignes Ecr <> aux montants Pièces/Lignes analytque ? ... }
        begin
            if ShuntPbAna and PourImport and PourNewImport and (ListePbAna <> nil) then
            begin
                PbAna := TRUE;
                ClePAF := ClePieceAnaFausse(MVT);
                ListePbAna.Add(ClePAF);
            end
            else
            begin
                if not MAJ then
                begin
                    LaListeAdd(GoListe1(Mvt, 'C', AxeEnCours + ' : ' + MsgLibel.Mess[15] + FloatToStr(Tot[3].TotDebit) + ' (' + FloatToStr(Tot[1].TotDebit) + ') ; '
                        + MsgLibel.Mess[16] + FloatToStr(Tot[3].TotCredit) + ' (' + FloatToStr(Tot[1].TotCredit) + ')'
                        , 3, 'A'));
                    Result := False;
                end
                else CorrEquilAna(2, Tot);
            end;
        end
        else
            if not CBonD then // B 13 { Devise : affichage dans la liste des Montants des Pièces/Lignes Ecr <> aux montants Pièces/Lignes analytque ? ... }
        begin
            if ShuntPbAna and PourImport and PourNewImport and (ListePbAna <> nil) then
            begin
                PbAna := TRUE;
                ClePAF := ClePieceAnaFausse(MVT);
                ListePbAna.Add(ClePAF);
            end
            else
            begin
                if not MAJ then
                begin
                    LaListeAdd(GoListe1(Mvt, 'C'
                        , AxeEnCours + ' : ' + MsgLibel.Mess[15] + FloatToStr(QAnaGene.FindField('SumDDev').AsFloat) + ' (' + FloatToStr(Tot[2].TotDebit) + ') ; '
                        + MsgLibel.Mess[16] + FloatToStr(QAnaGene.FindField('SumCDev').AsFloat) + ' (' + FloatToStr(Tot[2].TotCredit) + ')'
                        , 4, 'A'));
                    Result := False;
                end
                else CorrEquilAna(3, Tot);
            end;
        end;

        { Si la devise est renseigné les montants doivent etre <> 0 }
        if (Mvt.DEVISE <> V_PGI.DevisePivot) then
            if (Mvt.NATUREPIECE <> 'ECC') then // B 17
                if (QAnaGene.FindField('SumDDev').AsFloat = 0) and (QAnaGene.FindField('SumCDev').AsFloat = 0) then
                begin
                    if not MAJ then
                    begin
                        LaListeAdd(GoListe1(Mvt, 'C', AxeEnCours + ' : ' + ' (' + Mvt.DEVISE + ') ' + MsgLibel.Mess[15] + FloatToStr(QAnaGene.FindField('SumDDev').AsFloat)
                            + ' ; ' + MsgLibel.Mess[16] + FloatToStr(QAnaGene.FindField('SumCDev').AsFloat)
                            , 19, 'A'));
                        Result := False;
                    end
                    else CorrEquilAna(4, Tot);
                end;

        { Verifs autres infos d'Ana en rapport avec Ecr}

        VerifAutres(InfAna);
        if (InfAna.General <> Mvt.General) then {Meme Cpt Gene ?}
        begin // B 15
            { Cpt Gene Ecr Anal <> Cpt Gene Ecr , Affichage dans la lise }
            if not MAJ then
            begin
                LaListeAdd(GoListe1(Mvt, 'C', AxeEnCours + ' : ' + '(' + InfAna.General + ') ' + ' (' + Mvt.General + ')', 13, 'A'));
                Result := False;
            end
            else Corrige(QAnaGene, 'Y_GENERAL', Mvt.General); //MAJMVT('Y_GENERAL',Mvt.General) ;
        end;

        if InfAna.AnalPur then // B 16
        begin
            { Les Analytiques ne doivent pas être pur }
            if not MAJ then
            begin
                LaListeAdd(GoListe1(Mvt, 'C', '', 17, 'A'));
                Result := False;
            end
            else Corrige(QAnaGene, 'Y_TYPEANALYTIQUE', '-');
        end;
        QAnaGene.Next;
    end;
    (*
    else
    begin                     // B 11
        { L'écriture n'a pas d' analytiques rattachées, Affichage dans la liste }
        if Not MAJ then
        begin
            LaListeAdd(GoListe1(Mvt,'C',MsgLibel.Mess[37]+FloatToStr(Mvt.DEBIT)+' ; '+MsgLibel.Mess[38]+FloatToStr(Mvt.Credit),6,'A'));
            Result := False;
        end
        else
        begin
            Fillchar(Tot,SizeOf(Tot),#0);
            Tot[1].TotDebit := Arrondi(Mvt.DEBIT,V_PGI.OkdecV);
            Tot[1].TotCredit := Arrondi(Mvt.CREDIT,V_PGI.OkdecV);

            if Mvt.Devise<>V_PGI.DevisePivot then
            begin
                DecDev := 0;
                TrouveDecDev(Mvt.DEVISE,DecDev);
                Tot[2].TotDebit := Arrondi(Mvt.DEBITDEV,DecDev);
                Tot[2].TotCredit := Arrondi(Mvt.CREDITDEV,DecDev);
            end
            else Tot[2].TotDebit := Arrondi(Mvt.DEBITDEV,V_PGI.OkdecV);

            Tot[3].TotCredit := Arrondi(Mvt.CREDITDEV,V_PGI.OkdecV);

            Tot[3].TotDebit := 0;
            Tot[3].TotCredit := 0;
            Tot[4].TotDebit := 0;
            Tot[4].TotCredit := 0;

            CorrEquilAna(1, Tot);
        end;
    end;
    *)

    for i := 1 to 5 do
    begin
        if (i in SetAxe) and (AxeOk[i] = FALSE) then
        begin
            if not Maj then
            begin
                if ShuntPbAna then
                begin
                    PbAna := TRUE;
                    if PourImport and PourNewImport and (ListePbAna <> nil) then
                    begin
                        ClePAF := ClePieceAnaFausse(MVT);
                        ListePbAna.Add(ClePAF);
                    end;
                end
                else
                begin
                    AxeEnCours := 'A' + IntToStr(i);
                    LaListeAdd(GoListe1(Mvt, 'C', AxeEnCours + ' : ' + MsgLibel.Mess[37] + FloatToStr(Mvt.DEBIT) + ' ; ' + MsgLibel.Mess[38] + FloatToStr(Mvt.Credit), 6, 'A'));
                    Result := False;
                end;
            end
            else
            begin
                Fillchar(Tot, SizeOf(Tot), #0);
                Tot[1].TotDebit := Arrondi(Mvt.DEBIT, V_PGI.OkdecV);
                Tot[1].TotCredit := Arrondi(Mvt.CREDIT, V_PGI.OkdecV);

                if Mvt.Devise <> V_PGI.DevisePivot then
                begin
                    DecDev := 0;
                    TrouveDecDev(Mvt.DEVISE, DecDev);
                    Tot[2].TotDebit := Arrondi(Mvt.DEBITDEV, DecDev);
                    Tot[2].TotCredit := Arrondi(Mvt.CREDITDEV, DecDev);
                end
                else Tot[2].TotDebit := Arrondi(Mvt.DEBITDEV, V_PGI.OkdecV);

                Tot[3].TotCredit := Arrondi(Mvt.CREDITDEV, V_PGI.OkdecV);
                Tot[3].TotDebit := 0;
                Tot[3].TotCredit := 0;
                Tot[4].TotDebit := 0;
                Tot[4].TotCredit := 0;

                CorrEquilAna(1, Tot);
            end;
        end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.VerifAutres(var IFA : TINFOSMVTANA);
var
//    I : Byte;
    SQL : string;
begin
    Fillchar(IFA, SizeOf(IFA), #0);

//    if PourImport then I := 1
//    else I := 0;
//
//    QInfoAnaG.Params[0].AsString := Mvt.JOURNAL;
//    QInfoAnaG.Params[1].AsString := Mvt.EXERCICE;
//    QInfoAnaG.Params[2 - i].AsDateTime := Mvt.DATECOMPTABLE;
//    QInfoAnaG.Params[3 - i].AsInteger := Mvt.NUMEROPIECE;
//    QInfoAnaG.Params[4 - i].AsInteger := Mvt.NUMLIGNE;
//    QInfoAnaG.Params[5 - i].AsString := Mvt.DEVISE;
//    QInfoAnaG.Params[6 - i].AsString := Mvt.AXEPLUS[1];
//    QInfoAnaG.Params[7 - i].AsString := Mvt.AXEPLUS[2];
//    QInfoAnaG.Params[8 - i].AsString := Mvt.AXEPLUS[3];
//    QInfoAnaG.Params[9 - i].AsString := Mvt.AXEPLUS[4];
//    QInfoAnaG.Params[10 - i].AsString := Mvt.AXEPLUS[5];
//    QInfoAnaG.Params[11 - i].AsString := Mvt.ETABLISSEMENT;
//    QInfoAnaG.Params[12 - i].AsString := Mvt.QUALIFPIECE;

    if (not (QInfoAnaG = nil)) then Ferme(QInfoAnaG);

    SQL := ' Select ' + FY + '_GENERAL, ' + FY + '_TYPEANALYTIQUE, ' + FY + '_TYPEANOUVEAU, ' + FY + '_ECRANOUVEAU, ';
    SQL := SQL + ' ' + FY + '_JOURNAL, ' + FY + '_DATECOMPTABLE, ' + FY + '_NUMLIGNE, ';
    if not PourImport then SQL := SQL + ' Y_EXERCICE,  Y_NUMEROPIECE  '
    else SQL := SQL + ' IE_NUMPIECE, IE_CHRONO  ';
    SQL := SQL + ' From ' + FICANA + ' ';
    SQL := SQL + ' Where ';
    SQL := SQL + ' ' + FY + '_JOURNAL="' + Mvt.JOURNAL + '" ';
    if not PourImport then SQL := SQL + ' And Y_EXERCICE="' + Mvt.EXERCICE + '" ';
    SQL := SQL + ' And ' + FY + '_DATECOMPTABLE="' + UsDateTime(Mvt.DATECOMPTABLE) + '" ';
    if not PourImport then SQL := SQL + ' And Y_NUMEROPIECE="' + IntToStr(Mvt.NUMEROPIECE) + '" '
    else SQL := SQL + ' And IE_NUMPIECE="' + IntToStr(Mvt.NUMEROPIECE) + '" ';
    SQL := SQL + ' And ' + FY + '_NUMLIGNE="' + IntToStr(Mvt.NUMLIGNE) + '" ';
    SQL := SQL + ' And ' + FY + '_DEVISE="' + Mvt.DEVISE + '" ';
    SQL := SQL + ' And ' + FY + '_AXE IN ("' + Mvt.AXEPLUS[1] + '","' + Mvt.AXEPLUS[2] + '","' + Mvt.AXEPLUS[3] + '","' + Mvt.AXEPLUS[4] + '","' + Mvt.AXEPLUS[5] + '") ';
    SQL := SQL + ' And ' + FY + '_ETABLISSEMENT="' + Mvt.ETABLISSEMENT + '" ';
    SQL := SQL + ' And ' + FY + '_QUALIFPIECE="' + Mvt.QUALIFPIECE + '" ';
    SQL := SQL + ' And ' + FY + '_QUALIFPIECE<>"C" ';
    if PourImport then SQL := SQL + ' And IE_TYPEECR="A" and IE_SELECTED="X" and IE_OKCONTROLE<>"D" ';
    if not PourImport then SQL := SQL + ' order by Y_JOURNAL, Y_EXERCICE,  Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE '
    else SQL := SQL + ' order by IE_CHRONO ';

    QInfoAnaG := OpenSQL(SQL,true);

    with IFA do
    begin
        General := QInfoAnaG.Fields[0].AsString;
        AnalPur := (QInfoAnaG.Fields[1].AsString = 'X');
        TypeANO := QInfoAnaG.Fields[2].AsString;
        EcrANO := QInfoAnaG.Fields[3].AsString;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.EnregEquilA; { Enregistrement des valeurs de la 1ère requete sur la table ECRITURE... }
var
    I, J : Byte; { ...Pour la vérif de l'equilibre de ECR-ANA }
begin
    if PourImport then I := 1
    else I := 0;

    Mvt := TInfoMvt.Create;
    Mvt.JOURNAL := QSumGeAn.Fields[0].AsString; // journal
    Mvt.DATECOMPTABLE := QSumGeAn.Fields[2 - I].AsDateTime;
    Mvt.DEVISE := QSumGeAn.Fields[4 - I].AsString;
    Mvt.DEBIT := QSumGeAn.Fields[5 - I].AsFloat;
    Mvt.CREDIT := QSumGeAn.FieldS[6 - I].AsFloat;
    Mvt.DEBITDEV := QSumGeAn.FieldS[7 - I].AsFloat;
    Mvt.CREDITDEV := QSumGeAn.FieldS[8 - I].AsFloat;
    Mvt.NUMLIGNE := QSumGeAn.FieldS[9 - I].AsInteger;
    Mvt.REFINTERNE := QSumGeAn.Fields[10 - I].AsString;
    Mvt.GENERAL := QSumGeAn.Fields[11 - I].AsString;
    Mvt.ETABLISSEMENT := QSumGeAn.Fields[12 - I].AsString;
    Mvt.NATUREPIECE := QSumGeAn.Fields[13 - I].AsString;
    Mvt.ECRANOUVEAU := QSumGeAn.Fields[14 - I].AsString;

    for j := 1 to high(Mvt.AxePlus) do Mvt.AxePlus[j] := 'A0';

    if QSumGeAn.Fields[15 - I].AsString = 'X' then Mvt.AxePlus[1] := 'A1';
    if QSumGeAn.Fields[16 - I].AsString = 'X' then Mvt.AxePlus[2] := 'A2';
    if QSumGeAn.Fields[17 - I].AsString = 'X' then Mvt.AxePlus[3] := 'A3';
    if QSumGeAn.Fields[18 - I].AsString = 'X' then Mvt.AxePlus[4] := 'A4';
    if QSumGeAn.Fields[19 - I].AsString = 'X' then Mvt.AxePlus[5] := 'A5';

(*
    if QSumGeAn.Fields[17-I].AsString = 'X' then Mvt.AXE := 'A1';
    if QSumGeAn.Fields[18-I].AsString = 'X' then Mvt.AXE := 'A2';
    if QSumGeAn.Fields[19-I].AsString = 'X' then Mvt.AXE := 'A3';
    if QSumGeAn.Fields[20-I].AsString = 'X' then Mvt.AXE := 'A4';
    if QSumGeAn.Fields[21-I].AsString = 'X' then Mvt.AXE := 'A5';
*)

    Mvt.QUALIFPIECE := QSumGeAn.Fields[21 - I].AsString;

    if not PourImport then
    begin
        Mvt.NUMEROPIECE := QSumGeAn.Fields[3].AsInteger;
        Mvt.EXERCICE := QSumGeAn.FieldS[1].AsString;
    end
    else
    begin
        Mvt.NUMEROPIECE := QSumGeAn.FieldS[2].AsInteger;
        Mvt.EXERCICE := 'IMP';
        Mvt.CHRONO := QSumGeAn.Fields[21].AsInteger;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.CorrEquilAna(E : Byte;M : TabTot);
var
    T, Q : TQuery;
    St : string;
    NumVent : Integer;
    SolEP, SolED, SolAP, SolAD, MontantPivot, MontantDevise : Double;
begin
    if E <> 4 then
    begin
        St := ' Select Y_NUMVENTIL From ANALYTIQ';
        St := St + ' where Y_JOURNAL="' + Mvt.Journal + '" and Y_DATECOMPTABLE="' + USDATETIME(Mvt.DATECOMPTABLE) + '" ';
        St := St + ' and Y_EXERCICE="' + Mvt.Exercice + '" and Y_NUMEROPIECE=' + IntToStr(Mvt.NUMEROPIECE) + ' ';
        St := St + ' and Y_NUMLIGNE=' + IntToStr(Mvt.NUMLIGNE) + ' and Y_DEVISE="' + Mvt.DEVISE + '" ';
        St := St + ' and Y_ETABLISSEMENT="' + Mvt.ETABLISSEMENT + '" and Y_AXE="' + Mvt.Axe + '" AND Y_QUALIFPIECE="' + Mvt.QUALIFPIECE + '" ';
        St := St + ' order by Y_NUMVENTIL desc ';
        Q := OpenSql(St, True);
        NumVent := Q.FindField('Y_NUMVENTIL').AsInteger + 1;
        Ferme(Q);
    end
    else NumVent := 1;

    T := OpenSQL('SELECT * FROM ANALYTIQ WHERE Y_SECTION="Ed#"', False);
    T.Insert;
    InitNew(T);
    T.FindField('Y_JOURNAL').AsString := Mvt.JOURNAL;
    T.FindField('Y_EXERCICE').AsString := Mvt.EXERCICE;
    T.FindField('Y_DATECOMPTABLE').AsDateTime := Mvt.DATECOMPTABLE;
    T.FindField('Y_NUMEROPIECE').AsInteger := Mvt.NUMEROPIECE;
    T.FindField('Y_NUMLIGNE').AsInteger := Mvt.NUMLIGNE;
    T.FindField('Y_NUMVENTIL').AsInteger := NumVent;
    T.FindField('Y_DEVISE').AsString := Mvt.DEVISE;
    T.FindField('Y_REFINTERNE').AsString := Mvt.REFINTERNE;
    T.FindField('Y_GENERAL').AsString := Mvt.GENERAL;
    T.FindField('Y_NATUREPIECE').AsString := Mvt.NATUREPIECE;
    T.FindField('Y_QUALIFPIECE').AsString := Mvt.QUALIFPIECE;
    T.FindField('Y_ETABLISSEMENT').AsString := Mvt.ETABLISSEMENT;
    T.FindField('Y_AXE').AsString := Mvt.AXE;
    T.FindField('Y_SECTION').AsString := VH^.Cpta[AxeToFb(Mvt.Axe)].Attente;
    T.FindField('Y_LIBELLE').AsString := MsgLibel2.Mess[2];
    T.FindField('Y_ECRANOUVEAU').AsString := Mvt.ECRANOUVEAU;

    SolEP := LeSolde(M[1].TotDebit, M[1].TotCredit);
    SolED := LeSolde(M[2].TotDebit, M[2].TotCredit);
    SolAP := LeSolde(M[3].TotDebit, M[3].TotCredit);
    SolAD := LeSolde(M[4].TotDebit, M[4].TotCredit);

    case E of
        1 :
            begin
                MontantPivot := LeSolde(SolEP, SolAP);
                MontantDevise := LeSolde(SolED, SolAD);
                if MontantPivot < 0 then
                begin
                    T.FindField('Y_DEBIT').AsFloat := 0;
                    T.FindField('Y_CREDIT').AsFloat := Abs(MontantPivot);
                end
                else
                begin
                    T.FindField('Y_DEBIT').AsFloat := Abs(MontantPivot);
                    T.FindField('Y_CREDIT').AsFloat := 0;
                end;
                if MontantDevise < 0 then
                begin
                    T.FindField('Y_DEBITDEV').AsFloat := 0;
                    T.FindField('Y_CREDITDEV').AsFloat := Abs(MontantDevise);
                end
                else
                begin
                    T.FindField('Y_DEBITDEV').AsFloat := Abs(MontantDevise);
                    T.FindField('Y_CREDITDEV').AsFloat := 0;
                end;
            end;
        2 :
            begin
                MontantPivot := LeSolde(SolEP, SolAP);
                if MontantPivot < 0 then
                begin
                    T.FindField('Y_DEBIT').AsFloat := 0;
                    T.FindField('Y_CREDIT').AsFloat := Abs(MontantPivot);
                end
                else
                begin
                    T.FindField('Y_DEBIT').AsFloat := Abs(MontantPivot);
                    T.FindField('Y_CREDIT').AsFloat := 0;
                end;
            end;
        3 :
            begin
                MontantDevise := LeSolde(SolED, SolAD);
                if MontantDevise < 0 then
                begin
                    T.FindField('Y_DEBITDEV').AsFloat := 0;
                    T.FindField('Y_CREDITDEV').AsFloat := Abs(MontantDevise);
                end
                else
                begin
                    T.FindField('Y_DEBITDEV').AsFloat := Abs(MontantDevise);
                    T.FindField('Y_CREDITDEV').AsFloat := 0;
                end;
            end;
        4 :
            begin
                MontantPivot := LeSolde(SolEP, SolAP);
                if MontantPivot < 0 then
                begin
                    T.FindField('Y_DEBITDEV').AsFloat := 0;
                    T.FindField('Y_CREDITDEV').AsFloat := Abs(MontantPivot);
                end
                else
                begin
                    T.FindField('Y_DEBITDEV').AsFloat := Abs(MontantPivot);
                    T.FindField('Y_CREDITDEV').AsFloat := 0;
                end;
            end;
    end;

    T.Post;
    Ferme(T);
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.SQLLet; { Verif Lettrage : 1ere Partie.... }
begin
    QLettrage.first; { Si Etat Lettrage Ok }

    while not QLettrage.Eof do
    begin
        if Mvt <> nil then
        begin
            Mvt.FREE;
            Mvt := nil;
        end;

        EnregLet(QLettrage);
        MoveCur(False);

        if not VerifMvtNoLet then OkVerif := False;
        if not VerifMvtLet then OkVerif := False;

        TuFaisQuoi(QLettrage, 5);

        if not TestBreak then QLettrage.Next
        else Break;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.SQLVerifLet(Pourqui : Byte);
begin
    { Verif Lettrage : 2eme Partie...                    }
    { Pourqui=1 : Si E_LETTRAGEDEV Ok                    }
    { Pourqui=2 : Si Mountant couverture D C Ok          }
    {             sans Lettrage avec E_LETTRAGEDEV Error }
    Qletdev.First;

    while not QLetDev.Eof do
    begin
        MoveCur(False);
        case Pourqui of
            1 :
                begin
                    if not VerifLetDev then OkVerif := False;
                    if not TestBreak then QLetDev.Next
                    else Break;
                end;
            2 :
                begin
                    if OkPourSum then
                    begin
                        // suppresion de la fonction prepare .....
//                        PrepareLettrage;
                        if not SQLEquilLet then OkVerif := False;
                    end;
                    if not TestBreak then QLetDev.Next
                    else Break;
                end;
        end;
        TuFaisQuoi(QLetDev, 5);
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.SQLEquilLet : Boolean; { Appelé par OkPourSum }
var
    D, C, DDev, CDev : Double; { Verif si Montant des Couvertures au DEBIT }
    CBonP, CBonD : Boolean; { == Montant des Couvertures au CREDIT      }
    DecDev : Byte;
    SQL : string;
begin
    D := 0;
    C := 0;
    DDev := 0;
    CDev := 0;
    Result := True; // F 40

    if (not (QEquilLet = nil)) then Ferme(QEquilLet);

    SQL := ' Select E_AUXILIAIRE, E_GENERAL, E_LETTRAGE, E_ETATLETTRAGE, E_DEVISE, ';
    SQL := SQL + ' sum((E_couverture*(E_debit+E_DebitDev))/(E_debit+E_credit+E_DebitDev+E_CreditDev)) SumD, ';
    SQL := SQL + ' sum((E_couverture*(E_credit+E_CreditDev))/(E_debit+E_credit+E_DebitDev+E_CreditDev)) SumC, ';
    SQL := SQL + ' Sum((E_COUVERTUREDEV*(E_debit+E_DebitDev))/(E_debit+E_credit+E_DebitDev+E_CreditDev)) SumDDev,';
    SQL := SQL + ' Sum((E_COUVERTUREDEV*(E_credit+E_CreditDev))/(E_debit+E_credit+E_DebitDev+E_CreditDev)) SumCDev, ';
    SQL := SQL + ' E_LETTRAGEDEV, ';
    SQL := SQL + ' SUM(E_DEBIT + E_CREDIT + E_DEBITDEV + E_CREDITDEV) SUMT '; // 14747
    SQL := SQL + ' From ECRITURE ';
    SQL := SQL + ' Where ';
    SQL := SQL + ' E_QUALIFPIECE="N" ';
    SQL := SQL + ' AND E_AUXILIAIRE="' + QLetDev.Fields[1].AsString + '" ';
    SQL := SQL + ' AND E_GENERAL="' + QLetDev.Fields[2].AsString + '" ';
    SQL := SQL + ' AND E_LETTRAGE="' + QLetDev.Fields[0].AsString + '" ';
    if VH^.ExoV8.Code <> '' then SQL := SQL + ' And E_DATECOMPTABLE>="' + UsDateTime(VH^.ExoV8.Deb) + '" ';
    if VCrit.Etab <> '' then SQL := SQL + ' And E_ETABLISSEMENT="' + VCrit.Etab + '" ';
    SQL := SQL + ' And e_eche="X" and E_etatlettrage<>"" and E_etatlettrage<>"RI" ';
    SQL := SQL + ' And E_Ecranouveau<>"OAN" and E_ecranouveau<>"C" ';
    SQL := SQL + ' Group BY E_AUXILIAIRE, E_GENERAL, E_ETATLETTRAGE, E_LETTRAGE, E_DEVISE, E_LETTRAGEDEV ';

//    QEquilLet.Params[0].AsString := QLetDev.Fields[1].AsString;
//    QEquilLet.Params[1].AsString := QLetDev.Fields[2].AsString;
//    QEquilLet.Params[2].AsString := QLetDev.Fields[0].AsString;


    QEquilLet := OpenSQL(SQL,true);

    if QEquilLet.Fields[10].AsFloat = 0 then
      exit; // 14747

    while not QEquilLet.Eof do
    begin
        if Mvt <> nil then
        begin
            Mvt.FREE;
            Mvt := nil;
        end;

        EnregLet(QEquilLet);

        if QEquilLet.Fields[9].AsString = '-' then // E_LettrageDev
        begin
            D := Arrondi((D + QEquilLet.Fields[5].AsFloat), V_PGI.OkdecV); //SumD
            C := Arrondi((C + QEquilLet.Fields[6].AsFloat), V_PGI.OkdecV); //SumC

            if QEquilLet.Fields[4].AsString <> V_PGI.DevisePivot then //E_DEVISE
            begin
                TrouveDecDev(QEquilLet.Fields[4].AsString, DecDev); //E_DEVISE
                DDev := Arrondi((DDev + QEquilLet.Fields[7].AsFloat), DecDev); //SumDDev
                CDev := Arrondi((CDev + QEquilLet.Fields[8].AsFloat), DecDev); //SumCDev
            end
            else
            begin
                DDev := Arrondi((DDev + QEquilLet.Fields[7].AsFloat), V_PGI.OkdecV); //SumDDev
                CDev := Arrondi((CDev + QEquilLet.Fields[8].AsFloat), V_PGI.OkdecV); //SumCDev
            end;
        end
        else
            if QEquilLet.Fields[9].AsString = 'X' then // E_LettrageDev
        begin
            (* rony 9/04/97
            D:=Arrondi((D+QEquilLet.Fields[5].AsFloat),V_PGI.OkdecV) ;  //SumD
            C:=Arrondi((C+QEquilLet.Fields[6].AsFloat),V_PGI.OkdecV) ;  //SumC
            *)

            D := 0;
            C := 0;

            if QEquilLet.Fields[4].AsString <> V_PGI.DevisePivot then //E_DEVISE
            begin
                TrouveDecDev(QEquilLet.Fields[4].AsString, DecDev); //E_DEVISE
                DDev := Arrondi((DDev + QEquilLet.Fields[7].AsFloat), DecDev); //SumDDev
                CDev := Arrondi((CDev + QEquilLet.Fields[8].AsFloat), DecDev); //SumCDev
            end
            else
            begin
                DDev := Arrondi((DDev + QEquilLet.Fields[7].AsFloat), V_PGI.OkdecV); //SumDDev
                CDev := Arrondi((CDev + QEquilLet.Fields[8].AsFloat), V_PGI.OkdecV); //SumCDev
            end;
        end;
        QEquilLet.Next;
    end;

    CBonP := (D = C);
    CBonD := (DDev = CDev);

    if not CBonP then // F 41
    begin
        LaListeAdd(GoListe1(Mvt, 'L', MsgLibel.Mess[15] + ' ' + FloatToStr(D) + ' ; ' + MsgLibel.Mess[16] + ' ' + FloatToStr(C), 24, 'L'));
        Result := False;
    end;

    if not CBonD then // F 42 {Devise à vérifier, car SUM Dev avec E_lettrageDev=='X' !!}
    begin
        LaListeAdd(GoListe1(Mvt, 'L', MsgLibel.Mess[15] + ' ' + FloatToStr(DDev) + ' ; ' + MsgLibel.Mess[16] + ' ' + FloatToStr(CDev), 25, 'L'));
        Result := False;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.EnregLet(Q : TQuery);
begin
    Mvt := TInfoMvt.Create;
    Mvt.AUXILIAIRE := Q.FindField('E_AUXILIAIRE').AsString;
    Mvt.GENERAL := Q.FindField('E_GENERAL').AsString;
    Mvt.LETTRAGE := Q.FindField('E_LETTRAGE').AsString;
    Mvt.ETATLETTRAGE := Q.FindField('E_ETATLETTRAGE').AsString;

{$IFNDEF EAGL}
    if Q.Name = 'QLettrage' then
{$ELSE}
    if (Q = QLettrage) then
{$ENDIF}
    begin
        Mvt.DEBIT := Q.FindField('E_DEBIT').AsFloat;
        Mvt.CREDIT := Q.FindField('E_CREDIT').AsFloat;
        Mvt.DEBITDEV := Q.FindField('E_DEBITDEV').AsFloat;
        Mvt.CREDITDEV := Q.FindField('E_CREDITDEV').AsFloat;
        Mvt.COUVERTURE := Q.FindField('E_COUVERTURE').AsFloat;
        Mvt.COUVERTUREDEV := Q.FindField('E_COUVERTUREDEV').AsFloat;
        Mvt.DEVISE := Q.FindField('E_DEVISE').AsString;
        Mvt.LETTRAGEDEV := Q.FindField('E_LETTRAGEDEV').AsString;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.SQLEcrBud;
begin
    QEcrBud.First;

    while not QEcrBud.Eof do
    begin
        if MvtBud <> nil then
        begin
            MvtBud.FREE;
            MvtBud := nil;
        end;

        MoveCur(False);
        EnregMvtEcrBud;

        if VerifCompteBud then
        begin
            if VerifInfoJalBud then
            begin
                if VerifInfoExoBud then
                begin
                    if VerifInfoSecBud then
                    begin
                        if not VerifEtabBud then OkVerif := False;
                    end
                    else OkVerif := False;
                end
                else OkVerif := False;
            end
            else OkVerif := False;
        end
        else OkVerif := False;

        TuFaisQuoi(QEcrBud, 6);
        if TestBreak then Break;
        QEcrBud.Next;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.EnregMvtEcrBud;
begin {Pour Verif aussi avant imports}
    MvtBud := TInfoMvtBud.Create;
    MvtBud.DATECOMPTABLE := QEcrBud.Fields[0].AsDateTime;
    MvtBud.DEBIT := QEcrBud.Fields[1].AsFloat;
    MvtBud.CREDIT := QEcrBud.Fields[2].AsFloat;
    MvtBud.AXE := QEcrBud.Fields[3].AsString;
    MvtBud.REFINTERNE := QEcrBud.Fields[4].AsString;
    MvtBud.ETABLISSEMENT := QEcrBud.Fields[5].AsString;

    if not PourImport then
    begin
        MvtBud.EXERCICE := QEcrBud.Fields[6].AsString;
        MvtBud.NUMEROPIECE := QEcrBud.Fields[7].AsInteger;
        MvtBud.CONFIDENTIEL := QEcrBud.Fields[8].AsString;
        MvtBud.BUDJAL := QEcrBud.Fields[9].AsString;
        MvtBud.NATUREBUD := QEcrBud.Fields[10].AsString;
        MvtBud.BUDGENE := QEcrBud.Fields[11].AsString;
        MvtBud.BUDSECT := QEcrBud.Fields[12].AsString;
        MvtBud.QUALIFPIECE := QEcrBud.Fields[13].AsString;
    end
    else
    begin
        MvtBud.NUMEROPIECE := QEcrBud.Fields[6].AsInteger;
        MvtBud.CHRONO := QEcrBud.Fields[7].AsInteger;
        MvtBud.EXERCICE := 'IMP';
        MvtBud.BUDJAL := QEcrBud.Fields[8].AsString;
        MvtBud.NATUREBUD := QEcrBud.Fields[9].AsString;
        MvtBud.BUDGENE := QEcrBud.Fields[10].AsString;
        MvtBud.BUDSECT := QEcrBud.Fields[11].AsString;
        MvtBud.QUALIFPIECE := QEcrBud.Fields[12].AsString;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.LaListeAdd(G : DelInfo);
begin
    if (NbError = MaxNbErreur) then
        if (MsgRien.Execute(7, Caption, '') = mrCancel) then BStopClick(nil);
    if (NbError > MaxNbErreur) then Exit;
    LaListe.Add(G);
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.GoListe1(Entite : TObject;Quel, Rem : string;I : Byte;CodeIMP : string) : DelInfo;
var
    X : DelInfo;
    ImpError : TERRORIMP;
    StAxe : string;
    Entitee : TInfoMvt;
    Entitee1 : TInfoMvtBud;
    j : integer;
    ClePF : string;
begin
    Entitee := nil;
    Entitee1 := nil;
    if (Quel = 'B') then Entitee1 := TInfoMvtBud(Entite)
    else Entitee := TInfoMvt(Entite);

    Inc(NbError);
    X := DelInfo.Create;
    X.Gen := '';
    X.aux := '';
    X.Sect := '';
    X.GenACreer := FALSE;
    X.AuxACreer := FALSE;
    X.SectACreer := FALSE;

    if Quel = 'C' then
    begin
        X.Gen := Entitee.General;
        X.Aux := Entitee.Auxiliaire;
        X.Sect := Entitee.Section;
        if i = 11 then X.SectACreer := TRUE;
        if i = 10 then X.AuxACreer := TRUE;
        if i = 9 then X.GenACreer := TRUE;
        RetoucheDelInfo(X, Entitee);
        X.LeCod := Entitee.Journal;
        X.LeLib := Entitee.Refinterne;
        X.LeMess := IntToStr(Entitee.Numeropiece) + '/' + IntToStr(Entitee.Numligne);
        X.LeMess2 := DateToStr(Entitee.datecomptable);
        if i < 255 then X.LeMess3 := MsgLibel.Mess[i] + ' ' + Rem
        else X.LeMess3 := Rem;
        X.LeMess4 := Entitee.exercice;
        if (CodeIMP = 'A') then
        begin
            StAxe := '';
            for j := 1 to 5 do
                if Entitee.AXEPLUS[j] <> 'A0' then StAxe := StAxe + Entitee.AXEPLUS[j] + ';'
                else StAxe := StAxe + 'A1' + ';';
            X.LeMess4 := X.LeMess4 + ';' + StAxe;
        end
        else
            if (CodeIMP = 'Y') then X.LeMess4 := X.LeMess4 + ';' + Entitee.Axe;

        if Entitee.Axe <> '' then X.Axe := Entitee.Axe;
        // if (CodeIMP='A') or (CodeIMP='Y') then X.LeMess4:=Entitee.exercice+';'+Entitee.Axe else X.LeMess4:=Entitee.exercice ;
    end
    else
        if Quel = 'L' then
    begin
        X.LeCod := Entitee.lettrage;
        X.LeLib := Entitee.auxiliaire;
        X.LeMess := Entitee.general;
        X.LeMess2 := Entitee.ETATLETTRAGE;
        X.LeMess3 := MsgLibel.Mess[i] + ' ' + Rem;
        X.LeMess4 := Quel;
    end
    else
        if Quel = 'B' then
    begin
        X.LeCod := Entitee1.BudJal;
        X.LeLib := Entitee1.Refinterne;
        X.LeMess := IntToStr(Entitee1.Numeropiece);
        X.LeMess2 := DateToStr(Entitee1.datecomptable);
        if i < 255 then
            X.LeMess3 := MsgLibel.Mess[i] + ' ' + Rem
        else
            X.LeMess3 := Rem;
        X.LeMess4 := Entitee1.exercice;
        if (CodeIMP = 'A') then
        begin
            StAxe := '';
//            For j:=1 to 5 do StAxe:=StAxe+Entitee1.AXEPLUS[j]+';' ;
            for j := 1 to 5 do
                if Entitee.AXEPLUS[j] <> 'A0' then
                    StAxe := StAxe + Entitee1.AXEPLUS[j] + ';'
                else
                    StAxe := StAxe + 'A1' + ';';
            X.LeMess4 := X.LeMess4 + ';' + StAxe;
        end
        else
            if (CodeIMP = 'Y') then X.LeMess4 := X.LeMess4 + ';' + Entitee1.Axe;
        //if (CodeIMP='A') or (CodeIMP='Y') then X.LeMess4:=Entitee.exercice+';'+Entitee.Axe else X.LeMess4:=Entitee.exercice ;
    end;

    Result := X;
    if not PourImport then Exit;

    if PourNewImport then
    begin
        ClePF := ClePieceFausse(X);
        ImpListe.Add(ClePF);
    end
    else
    begin
        if ImpListe.Count <> 0 then { Pour Fichier Import}
        begin
            ImpError := TERRORIMP(ImpListe.Objects[ImpListe.Count - 1]);
            if Quel = 'B' then
            begin
                if ImpError.NumChrono = Entitee1.Chrono then exit;
            end
            else
            begin
                if ImpError.NumChrono = Entitee.Chrono then exit;
            end;
        end;
        ImpError := TERRORIMP.Create;
        if Quel = 'B' then
            ImpError.NumChrono := Entitee1.Chrono
        else
            ImpError.NumChrono := Entitee.Chrono;
        ImpError.NumErreur := CodeIMP;
        ImpListe.AddObject('', ImpError);
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.GoListe2(Entitee : TInfoMvt;Quel, Rem : string;F : TField;CodeIMP : string) : DelInfo;
var
    X : DelInfo;
    ImpError : TERRORIMP;
    ClePF : string;
begin
    Inc(NbError);
    X := DelInfo.Create;
    X.Gen := '';
    X.aux := '';
    X.Sect := '';
    X.GenACreer := FALSE;
    X.AuxACreer := FALSE;
    X.SectACreer := FALSE;

    if Quel = 'C' then
    begin
        X.Gen := Entitee.General;
        X.Aux := Entitee.Auxiliaire;
        X.Sect := Entitee.Section;
        RetoucheDelInfo(X, Entitee);
        X.LeCod := Entitee.Journal;
        X.LeLib := Entitee.Refinterne;
        X.LeMess := IntToStr(Entitee.Numeropiece) + '/' + IntToStr(Entitee.Numligne);
        X.LeMess2 := DateToStr(Entitee.datecomptable);
        X.LeMess3 := MsgLibel2.Mess[0] + ' "' + F.FieldName + '" ' + MsgLibel2.Mess[1] + ' ' + Rem;
        if (CodeIMP = 'A') or (CodeIMP = 'Y') then
            X.LeMess4 := Entitee.exercice + ';' + Entitee.Axe
        else
            X.LeMess4 := Entitee.exercice;
    end
    else
        if Quel = 'L' then
    begin
        X.LeCod := Entitee.lettrage;
        X.LeLib := Entitee.auxiliaire;
        X.LeMess := Entitee.general;
        X.LeMess2 := Entitee.ETATLETTRAGE;
        X.LeMess3 := MsgLibel2.Mess[0] + ' "' + F.FieldName + '" ' + MsgLibel2.Mess[1] + ' ' + Rem;
        X.LeMess4 := Quel;
    end;

    Result := X;
    if not PourImport then Exit;

    if PourNEwImport then
    begin
        ClePF := ClePieceFausse(X);
        ImpListe.Add(ClePF);
    end
    else
    begin
        if ImpListe.Count <> 0 then
        begin
            ImpError := TERRORIMP(ImpListe.Objects[ImpListe.Count - 1]);
            if ImpError.NumChrono = Entitee.Chrono then exit;
        end;
        ImpError := TERRORIMP.Create;
        ImpError.NumChrono := Entitee.Chrono;
        ImpError.NumErreur := CodeIMP;
        ImpListe.AddObject('', ImpError);
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.Majuscule(St : string) : Boolean;
begin
    Majuscule := (St = UpperCase(St));
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.TuFaisQuoi(Q : TQuery;DeQuoi : Byte);
begin
    if PourImport then exit;

    if Q <> nil then
    begin
        case Dequoi of
            4 :
                TTravail.Caption := MsgBar.Mess[1] + ' ' + Q.FindField('Y_JOURNAL').AsString + ' ' +
                    MsgBar.Mess[2] + ' ' + DateToStr(Q.FindField('Y_DATECOMPTABLE').AsDateTime) + ' ' +
                    MsgBar.Mess[3] + ' ' + IntToStr(Q.FindField('Y_NUMEROPIECE').AsInteger);
            5 :
                TTravail.Caption := MsgBar.Mess[11] + ' ' + Q.FindField('E_LETTRAGE').AsString + ' ' +
                    MsgBar.Mess[12] + ' ' + Q.FindField('E_AUXILIAIRE').AsString + ' ' +
                    MsgBar.Mess[13] + ' ' + Q.FindField('E_GENERAL').AsString;
            6 :
                TTravail.Caption := MsgBar.Mess[1] + ' ' + Q.FindField('BE_BUDJAL').AsString + ' ' +
                    MsgBar.Mess[2] + ' ' + DateToStr(Q.FindField('BE_DATECOMPTABLE').AsDateTime) + ' ' +
                    MsgBar.Mess[3] + ' ' + Q.FindField('BE_NUMEROPIECE').AsString;
            else
                TTravail.Caption := MsgBar.Mess[1] + ' ' + Q.FindField('E_JOURNAL').AsString + ' ' +
                    MsgBar.Mess[2] + ' ' + DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime) + ' ' +
                    MsgBar.Mess[3] + ' ' + IntToStr(Q.FindField('E_NUMEROPIECE').AsInteger);
        end;
    end;

    case DeQuoi of
        1 :
            if MAJ then
                TNBError1.Caption := MsgBar.Mess[7]
            else
            begin
                ErrEcr := NbError;
                if ErrEcr = 0 then TNBError1.Caption := MsgBar.Mess[7] + '... ' + MsgBar.Mess[6];
                if ErrEcr = 1 then TNBError1.Caption := MsgBar.Mess[7] + '... ' + IntToStr(ErrEcr) + ' ' + MsgBar.Mess[5];
                if ErrEcr > 1 then TNBError1.Caption := MsgBar.Mess[7] + '... ' + IntToStr(ErrEcr) + ' ' + MsgBar.Mess[4];
                if ErrEcr >= 1 then TNBError1.Font.color := clRed;
            end;
        2 :
            if MAJ then
                TNBError2.Caption := MsgBar.Mess[8]
            else
            begin
                ErrAnal := NbError - ErrEcr;
                if ErrAnal = 0 then TNBError2.Caption := MsgBar.Mess[8] + '... ' + MsgBar.Mess[6];
                if ErrAnal = 1 then TNBError2.Caption := MsgBar.Mess[8] + '... ' + IntToStr(ErrAnal) + ' ' + MsgBar.Mess[5];
                if ErrAnal > 1 then TNBError2.Caption := MsgBar.Mess[8] + '... ' + IntToStr(ErrAnal) + ' ' + MsgBar.Mess[4];
                if ErrAnal >= 1 then TNBError2.Font.color := clRed;
            end;
        3 :
            if MAJ then
                TNBError3.Caption := MsgBar.Mess[9]
            else
            begin
                ErrMvt := NbError - ErrEcr - ErrAnal;
                if ErrMvt = 0 then TNBError3.Caption := MsgBar.Mess[9] + '... ' + MsgBar.Mess[6];
                if ErrMvt = 1 then TNBError3.Caption := MsgBar.Mess[9] + '... ' + IntToStr(ErrMvt) + ' ' + MsgBar.Mess[5];
                if ErrMvt > 1 then TNBError3.Caption := MsgBar.Mess[9] + '... ' + IntToStr(ErrMvt) + ' ' + MsgBar.Mess[4];
                if ErrMvt >= 1 then TNBError3.Font.color := clRed;
            end;
        4 :
            if MAJ then
                TNBError4.Caption := MsgBar.Mess[15]
            else
            begin
                ErrMvtA := NbError - ErrEcr - ErrAnal - ErrMvt;
                if ErrMvtA = 0 then TNBError4.Caption := MsgBar.Mess[15] + '... ' + MsgBar.Mess[6];
                if ErrMvtA = 1 then TNBError4.Caption := MsgBar.Mess[15] + '... ' + IntToStr(ErrMvtA) + ' ' + MsgBar.Mess[5];
                if ErrMvtA > 1 then TNBError4.Caption := MsgBar.Mess[15] + '... ' + IntToStr(ErrMvtA) + ' ' + MsgBar.Mess[4];
                if ErrMvtA >= 1 then TNBError4.Font.color := clRed;
            end;
        5 :
            if FVerification.ItemIndex = 2 then
            begin
                if MAJ then
                    TNBError1.Caption := MsgBar.Mess[10]
                else
                begin
                    ErrLet := NbError;
                    if ErrLet = 0 then TNBError1.Caption := MsgBar.Mess[10] + '... ' + MsgBar.Mess[6];
                    if ErrLet = 1 then TNBError1.Caption := MsgBar.Mess[10] + '... ' + IntToStr(ErrLet) + ' ' + MsgBar.Mess[5];
                    if ErrLet > 1 then TNBError1.Caption := MsgBar.Mess[10] + '... ' + IntToStr(ErrLet) + ' ' + MsgBar.Mess[4];
                    if ErrLet >= 1 then TNBError1.Font.color := clRed;
                end;
            end
            else
            begin
                if MAJ then
                    TNBError1.Caption := MsgBar.Mess[10]
                else
                begin
                    ErrLet := NbError - ErrEcr - ErrAnal - ErrMvt - ErrMvtA;
                    if ErrLet = 0 then TNBError5.Caption := MsgBar.Mess[10] + '... ' + MsgBar.Mess[6];
                    if ErrLet = 1 then TNBError5.Caption := MsgBar.Mess[10] + '... ' + IntToStr(ErrLet) + ' ' + MsgBar.Mess[5];
                    if ErrLet > 1 then TNBError5.Caption := MsgBar.Mess[10] + '... ' + IntToStr(ErrLet) + ' ' + MsgBar.Mess[4];
                    if ErrLet >= 1 then TNBError5.Font.color := clRed;
                end;
            end;
        6 :
            if FVerification.ItemIndex = 3 then
            begin
                if MAJ then
                    TNBError1.Caption := MsgBar.Mess[22]
                else
                begin
                    ErrMvtB := NbError;
                    if ErrMvtB = 0 then TNBError1.Caption := MsgBar.Mess[22] + '... ' + MsgBar.Mess[6];
                    if ErrMvtB = 1 then TNBError1.Caption := MsgBar.Mess[22] + '... ' + IntToStr(ErrMvtB) + ' ' + MsgBar.Mess[5];
                    if ErrMvtB > 1 then TNBError1.Caption := MsgBar.Mess[22] + '... ' + IntToStr(ErrMvtB) + ' ' + MsgBar.Mess[4];
                    if ErrMvtB >= 1 then TNBError1.Font.color := clRed;
                end;
            end
            else
            begin
                if MAJ then
                    TNBError6.Caption := MsgBar.Mess[22]
                else
                begin
                    ErrMvtB := NbError - ErrEcr - ErrAnal - ErrMvt - ErrMvtA - ErrLet;
                    if ErrMvtB = 0 then TNBError6.Caption := MsgBar.Mess[22] + '... ' + MsgBar.Mess[6];
                    if ErrMvtB = 1 then TNBError6.Caption := MsgBar.Mess[22] + '... ' + IntToStr(ErrMvtB) + ' ' + MsgBar.Mess[5];
                    if ErrMvtB > 1 then TNBError6.Caption := MsgBar.Mess[22] + '... ' + IntToStr(ErrMvtB) + ' ' + MsgBar.Mess[4];
                    if ErrMvtB >= 1 then TNBError6.Font.color := clRed;
                end;
            end;
    end;
end;

//********** Fonction test D'EXISTANCE

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.Existe(V : Byte;St : string) : Boolean;
var
    i : Integer;
    Inf : TEnregInfos;
    trouve : Boolean;
begin
    Trouve := False;
    case V of
        1 :
            begin // Devise
                for i := 0 to ListeDEV.Count - 1 do
                begin
                    if ListeDEV.Objects[i] <> nil then
                    begin
                        Inf := TEnregInfos(ListeDEV.Objects[i]);
                        Trouve := (Inf.Inf1 = St);
                        if Trouve then Break;
                    end
                    else Break;
                end;
            end;
        2 :
            begin // Journal
                for i := 0 to ListeJAL.Count - 1 do
                begin
                    if ListeJAL.Objects[i] <> nil then
                    begin
                        Inf := TEnregInfos(ListeJAL.Objects[i]);
                        Trouve := (Inf.Inf1 = St);
                        if Trouve then
                        begin
                            InfAutres.J_NATUREJAL := Inf.Inf2;
                            InfAutres.J_Axe := Inf.Inf3;
                            InfAutres.J_COMPTEINTERDIT := Inf.Inf4;
                            InfAutres.J_MODESAISIE := Inf.Inf5;
                            InfAutres.J_VALIDEEN := Inf.Inf6; // FQ 12405
                            InfAutres.J_VALIDEEN1 := Inf.Inf7; // FQ 12405
                            Break;
                        end;
                    end
                    else Break;
                end;
            end;
        3 :
            begin // Nature de Piece
                Trouve := (TCNatPiece.Items[TCNatPiece.Items.IndexOf(St)] = St);
            end;
    end;
    Result := Trouve;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.OkExo(E : String3;D : TDateTime) : Boolean;
var
    i : Integer;
    Inf : TEnregInfos;
    Ok : Boolean;
begin
    Ok := False;

    for i := 0 to ListeEXO.Count - 1 do
    begin
        if ListeEXO.Objects[i] <> nil then
        begin
            Inf := TEnregInfos(ListeEXO.Objects[i]);
            if (Inf.Inf1 = E) then
            begin
                if (D >= Inf.D1) then
                    if (D <= Inf.D2) then
                    begin
                        InfAutres.EX_DATEEXO[0] := Inf.D1;
                        InfAutres.EX_DATEEXO[1] := Inf.D2;
                        Ok := True;
                    end;
                Break;
            end;
        end
        else Break;
    end;

    Result := Ok;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFVerCpta.TrouveDecDev(St : String3;var Dec : Byte);
var
    i : Integer;
    Inf : TEnregInfos;
begin
    for i := 0 to ListeDEV.Count - 1 do
    begin
        if ListeDEV.Objects[i] <> nil then
        begin
            Inf := TEnregInfos(ListeDEV.Objects[i]);
            if (Inf.Inf1 = St) then
            begin
                Dec := Inf.Nb;
                Break;
            end;
        end
        else Break;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.ExisteSec(Ax, Cod : string) : Boolean;
var
    i : Integer;
    Inf : TEnregInfos;
    Trouve : Boolean;
begin
    Trouve := False;

    for i := 0 to ListeSec.Count - 1 do
    begin
        if ListeSec.Objects[i] <> nil then
        begin
            Inf := TEnregInfos(ListeSec.Objects[i]);
            Trouve := ((Inf.Inf1 = Ax) and (Inf.Inf2 = Cod));

            if Trouve then Break;
        end
        else Break;
    end;

    Result := Trouve;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.OkPourSum : Boolean; { Appelé par SQLVerifLet(2) }
var
    Trouve : Boolean;
    i : Integer;
    ErLetDev : TErrLetDev;
begin
    Result := True;

    for i := 0 to LetListe.Count - 1 do
    begin
        if LetListe.Objects[i] <> nil then
        begin
            ErLetDev := TErrLetDev(LetListe.Objects[i]);

            Trouve := ((ErLetDev.Lettres = QLetDev.FindField('E_LETTRAGE').AsString) and (ErLetDev.Auxi = QLetDev.FindField('E_AUXILIAIRE').AsString) and (ErLetDev.Gene = QLetDev.FindField('E_GENERAL').AsString) and (ErLetDev.EtatLet = QLetDev.FindField('E_ETATLETTRAGE').AsString));

            if Trouve then
            begin
                Break;
                Result := False;
            end;
        end
        else Break;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.ExisteCpteBud(cod : string) : boolean;
var
    i : Integer;
    Inf : TEnregInfos;
    Trouve : Boolean;
begin
    Trouve := False;

    for i := 0 to ListeCptBud.Count - 1 do
    begin
        if ListeCptBud.Objects[i] <> nil then
        begin
            Inf := TEnregInfos(ListeCptBud.Objects[i]);
            Trouve := (Inf.Inf1 = Cod);

            if Trouve then Break;
        end
        else Break;
    end;

    Result := Trouve;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.ExisteJalBud(cod : string) : boolean;
var
    i : Integer;
    Inf : TEnregInfos;
    Trouve : Boolean;
begin
    Trouve := False;

    for i := 0 to ListeJalBud.Count - 1 do
    begin
        if ListeJalBud.Objects[i] <> nil then
        begin
            Inf := TEnregInfos(ListeJalBud.Objects[i]);

            Trouve := (Inf.Inf1 = Cod);
            if Trouve then Break;
        end
        else Break;
    end;

    Result := Trouve;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.ExisteSecBud(Ax, Cod : string) : Boolean;
var
    i : Integer;
    Inf : TEnregInfos;
    Trouve : Boolean;
begin
    Trouve := False;

    for i := 0 to ListeSecBud.Count - 1 do
    begin
        if ListeSecBud.Objects[i] <> nil then
        begin
            Inf := TEnregInfos(ListeSecBud.Objects[i]);
            Trouve := ((Inf.Inf1 = Ax) and (Inf.Inf2 = Cod));

            if Trouve then Break;
        end
        else Break;
    end;

    Result := Trouve;
end;

//********** Fonction de VERIFICATION

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.VerifCompte(var let, Poin, vent : Boolean) : Boolean; // C 10
{ Verif si le Cpt existe ou s'il doit etre renseigné }
begin
    Result := True;
    Let := False;
    Poin := False;
    Vent := False;

    { Verif si Cpt Géné et Tiers est renseigné ou existe }
    if Mvt.General = '' then // C 11
    begin
        LaListeAdd(GoListe1(Mvt, 'C', ' "' + Mvt.General + '" ', 57, 'E'));
        VerifCompte := False;
        Exit;
    end;

    Fillchar(CGen, SizeOf(CGen), #0);
    InfosCptGen(Mvt.General);
    Fillchar(CTiers, SizeOf(CTiers), #0);
    InfosCptAux(Mvt.Auxiliaire);

    if Cgen.General <> '' then
    begin
        if CGen.Collectif then { Si Gene est collectif }
        begin
            if Mvt.AUXILIAIRE <> '' then { Si Tiers Renseigné }
            begin
                if CTiers.Auxi = '' then { Si le Tiers n'existe pas } // C 16
                begin {  }
                    LaListeAdd(GoListe1(Mvt, 'C', '(' + Mvt.AUXILIAIRE + ') ', 10, 'E'));
                    Result := False;
                end;
                Let := CTiers.Lettrable; {Modif : 24/12}
            end
            else { Si Tiers Non Renseigné }
            begin
                { Sinon, Le Tiers doit etre renseigné }// C 14
                if MAJ then Corrige(QEcr, 'E_AUXILIAIRE', 'RONY') { A Supp, bien sur ...}
                else
                begin
                    LaListeAdd(GoListe1(Mvt, 'C', '', 7, 'E'));
                    Result := False;
                end;
            end;
        end
        else { Si Gene est Non - collectif }
        begin
            { Peut-etre parce que Le Géné est TID ou TIC }
            if (CGen.NatureGene = 'TID') or (CGen.NatureGene = 'TIC') then
            begin
                Let := CGen.Lettrable; {Modif : 24/12}
            end;
            if Mvt.AUXILIAIRE <> '' then
            begin { Le Tiers doit etre vide } // C 15
                LaListeAdd(GoListe1(Mvt, 'C', '', 8, 'E'));
                Result := False;
            end;
            Poin := CGen.Pointable; {Modif : 24/12}
        end;
(*
        If CGen.NatureGene = 'BQE' then   // C 13
        begin
            DevBqeLu:=DevBqe(CGen.General);
            if DevBqe(CGen.General)<>Mvt.DEVISE then
            begin
                If DevBqeLu = '' Then LaListeAdd(GoListe1(Mvt,'C','('+CGen.General+') '+'('+Mvt.DEVISE+') ',94,'E'))
                Else LaListeAdd(GoListe1(Mvt,'C','('+CGen.General+') '+'('+Mvt.DEVISE+') ',33,'E'));
                Result := False;
            end;
        end;
*)
        vent := CGen.Ventilable[1] or CGen.Ventilable[2] or CGen.Ventilable[3] or CGen.Ventilable[4] or CGen.Ventilable[5];
    end
    else // C 12
    begin
        LaListeAdd(GoListe1(Mvt, 'C', '(' + Mvt.GENERAL + ') ', 9, 'E'));
        Result := False;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.VerifInfoJal : Boolean;
var
    Sens : Byte;
//    wYear, wMonth, wDay : Word; // FQ 12405
begin
    Result := True;
    Sens := 0;
    Fillchar(InfAutres, SizeOf(InfAutres), #0);

    if Mvt.JOURNAL <> '' then
    begin // C 12
        if not Existe(2, Mvt.JOURNAL) then
        begin
            LaListeAdd(GoListe1(Mvt, 'C', ' "' + Mvt.JOURNAL + '" ' + MsgLibel.Mess[37] + FloatToStr(Mvt.DEBIT) + ' ; ' + MsgLibel.Mess[38] + FloatToStr(Mvt.CREDIT), 12, 'E'));
            Result := False;
        end
        else
        begin
            if Mvt.DEVISE = V_PGI.DevisePivot then
            begin
                if Mvt.DEBIT <> 0 then Sens := 1
                else
                    if Mvt.CREDIT <> 0 then Sens := 2;
            end
            else
            begin
                if Mvt.DEBITDEV <> 0 then Sens := 1
                else
                    if Mvt.CREDITDEV <> 0 then Sens := 2;
            end;

            if Sens <> 0 then
                if EstInterdit(InfAutres.J_COMPTEINTERDIT, Mvt.GENERAL, Sens) <> 0 then
                begin // C 13
                    LaListeAdd(GoListe1(Mvt, 'C', ' (' + Mvt.GENERAL + ') ', 34, 'E'));
                    Result := False;
                end;

            if InfAutres.J_NATUREJAL = 'BQE' then
            begin
                if (Mvt.NaturePiece <> 'RF') and (Mvt.NaturePiece <> 'RC') and (Mvt.NaturePiece <> 'OD') and
                    (Mvt.NaturePiece <> 'OF') and (Mvt.NaturePiece <> 'OC') then
                begin // C ??
                    LaListeAdd(GoListe1(Mvt, 'C', ' (' + Mvt.Journal + ' / ' + Mvt.NaturePiece + ') ', 105, 'E'));
                    Result := False;
                end;
            end;

            if InfAutres.J_MODESAISIE <> Mvt.ModeSaisie then
            begin
                if ((InfAutres.J_MODESAISIE = '') and (Mvt.ModeSaisie = '-')) or
                    ((InfAutres.J_MODESAISIE = '-') and (Mvt.ModeSaisie = '')) then
                else
                begin
                    LaListeAdd(GoListe1(Mvt, 'C', ' (' + Mvt.Journal + ' / ' + Mvt.NaturePiece + ') ', 111, 'E'));
                    Result := False;
                end;
            end;
        end;

// BPY le 15/11/2004 => Fiche n° 14974 : pb lors de l'import .. on doit pouvoir importé sur une periode validé !
//        // FQ 12405  - Interdit d'importer des écritures sur une période clôturée
//        Decodedate(Mvt.DATECOMPTABLE, wYear, wMonth, wDay);
//
//        if (InfAutres.J_VALIDEEN <> '') and (InfAutres.J_VALIDEEN1 <> '') and PourImport then
//        begin // 13308
//            if (InfAutres.J_VALIDEEN[wMonth] = 'X') and (VH^.EnCours.Code = QUELEXO(DateToStr(Mvt.DATECOMPTABLE))) or (InfAutres.J_VALIDEEN1[wMonth] = 'X') and (VH^.Suivant.Code = QUELEXO(DateToStr(Mvt.DATECOMPTABLE))) then
//            begin
//                LaListeAdd(GoListe1(Mvt, 'C', ' (' + Mvt.Journal + ' / ' + DateToStr(Mvt.DATECOMPTABLE) + ') ', 114, 'E'));
//                Result := False;
//            end;
//        end;
    end
    else
    begin // C 11
        LaListeAdd(GoListe1(Mvt, 'C', '', 46, 'E'));
        Result := False;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 15/11/2004
Modifié le ... : 15/11/2004
Description .. : - BPY le 15/11/2004 - ajout d'un test pour la cloture 
Suite ........ : perodique 
Mots clefs ... : 
*****************************************************************}

function TFVerCpta.VerifInfoPerio : Boolean;
begin
    result := true;

    if PourImport then
    begin
        if (Mvt.DATECOMPTABLE <= VH^.DateCloturePer) then
        begin
            LaListeAdd(GoListe1(Mvt, 'C', ' (' + DateToStr(Mvt.DATECOMPTABLE) + ') ', 114, 'E'));
            Result := False;
        end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.VerifInfoExo : Boolean;
var
    Exo : TExoDate;
    MonoExo : Boolean;
begin
    Result := True; //E_EXERCICE Error   // C90
    if PourImport then
    begin
(*
        InfAutres.EX_DATEEXO[0] := VH^.EnCours.Deb;
        InfAutres.EX_DATEEXO[1] := VH^.EnCours.Fin;

        if (Mvt.DATECOMPTABLE>=VH^.Precedent.Deb) and (Mvt.DATECOMPTABLE<=VH^.Precedent.Fin) and (VH^.Precedent.Code<>'') then
        begin
            InfAutres.EX_DATEEXO[0] := VH^.Precedent.Deb;
            InfAutres.EX_DATEEXO[1] := VH^.Precedent.Fin;
        end
        else if (Mvt.DATECOMPTABLE>=VH^.EnCours.Deb) and (Mvt.DATECOMPTABLE<=VH^.EnCours.Fin) then
        begin
            InfAutres.EX_DATEEXO[0] := VH^.EnCours.Deb;
            InfAutres.EX_DATEEXO[1] := VH^.EnCours.Fin;
        end
        else if (Mvt.DATECOMPTABLE>=VH^.Suivant.Deb) and (Mvt.DATECOMPTABLE<=VH^.Suivant.Fin) and (VH^.Suivant.Code<>'') then
        begin
            InfAutres.EX_DATEEXO[0] := VH^.Suivant.Deb;
            InfAutres.EX_DATEEXO[1] := VH^.Suivant.Fin;
        end;
*)
        InfAutres.EX_DATEEXO[0] := VH^.EnCours.Deb;
        InfAutres.EX_DATEEXO[1] := VH^.EnCours.Fin;
        if not QuelExoDate(Mvt.DATECOMPTABLE, Mvt.DATECOMPTABLE, MonoExo, Exo) then
        begin
            Result := FALSE;
            LaListeAdd(GoListe1(Mvt, 'C', '', 104, 'E'));
        end
        else
        begin
            if (not RecupSISCO) and (not RecupServant) and (not RecupSISCOExt) then
            begin
                if (Exo.code <> VH^.EnCours.Code) and (Exo.code <> VH^.Suivant.Code) then
                begin
                    Result := FALSE;
                    LaListeAdd(GoListe1(Mvt, 'C', '', 104, 'E'));
                end;
            end;
            InfAutres.EX_DATEEXO[0] := Exo.Deb;
            InfAutres.EX_DATEEXO[1] := Exo.Fin;
        end;
        exit;
    end;

    if not OkExo(Mvt.EXERCICE, Mvt.DATECOMPTABLE) then
    begin
        LaListeAdd(GoListe2(Mvt, 'C', ' (' + Mvt.Exercice + ') ', QEcr.FindField('E_EXERCICE'), 'E'));
        Result := False;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.VerifInfoEcrAno : Boolean;
begin
    Result := True;

    if not EcrAnoOk(InfAutres.J_NATUREJAL, Mvt.ECRANOUVEAU) then
    begin // C 31
        LaListeAdd(GoListe2(Mvt, 'C', ' (' + Mvt.ECRANOUVEAU + ') ', QEcr.FindField('' + FE + '_ECRANOUVEAU'), 'E'));
        Result := False;
    end;

    if InfAutres.J_NATUREJAL = 'ANO' then
    begin
        if (CGen.Naturegene = 'CHA') or (CGen.NatureGene = 'PRO') then
        begin // C 32
            LaListeAdd(GoListe1(Mvt, 'C', '(' + Mvt.General + ' )', 29, 'E'));
            Result := False;
        end;

        if (CGen.Naturegene = 'EXT') then // xdoc
        begin
            LaListeAdd(GoListe1(Mvt, 'C', '(' + Mvt.General + ' )', 64, 'E'));
            Result := False;
        end;

        if Mvt.DateComptable <> InfAutres.EX_DATEEXO[0] then // xdoc
        begin
            LaListeAdd(GoListe1(Mvt, 'C', '(' + DateToStr(InfAutres.EX_DATEEXO[0]) + ' )', 67, 'E'));
            Result := False;
        end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.VerifInfoEche(Lettrable, pointable : Boolean) : Boolean;
var
    OkEche, OkModP, OkPourLet, OkPourPoint : Boolean;
    Oknumeche : Boolean;
begin
    Result := True;
    //Ok:=True ;
    OkModP := True;
    Oknumeche := True;
    OkEche := True;
    OkPourLet := (Mvt.EtatLettrage = 'AL') or (Mvt.EtatLettrage = 'PL') or (Mvt.EtatLettrage = 'TL');
    OkPourPoint := (Mvt.EtatLettrage = 'RI');

    if lettrable then
    begin // C 40
        if OkPourLet then {  }
        begin
            //Ok:=False ; {  OK = Ok pour e_eche }
            //OkNumeche=???}
            if Mvt.ECHE = 'X' then
            begin
                //Ok:=False ;
                //OkModP:=False ;
                if (Mvt.MODEPAIE <> '') and (TCModP.values.IndexOf(Mvt.MODEPAIE) > -1) then
                begin
                    OkModP := True;
                    if Mvt.NUMECHE < 1 then Oknumeche := False;
                end
                else OkModP := False;
            end
            else OkEche := False;
        end
        else
        begin
            LaListeAdd(GoListe1(Mvt, 'C', ' (' + Mvt.EtatLettrage + ') ', 91, 'E'));
            Result := False;
            Exit;
        end;
        if not OkEche then
        begin { Msg E_ECHE Error }
            LaListeAdd(GoListe2(Mvt, 'C', '', QEcr.FindField('' + FE + '_ECHE'), 'E'));
            Result := False;
            //OK:=False ;
        end;

        if not Oknumeche then
        begin { Msg E_NUMECHE Error }
            LaListeAdd(GoListe2(Mvt, 'C', '', QEcr.FindField('' + FE + '_NUMECHE'), 'E'));
            Result := False;
            //OK:=False ;
        end;
{   If Not OkEche Then ...}
        if not OkModP then { Msg E_MODEPAIE vide ou inconnue }
        begin
            if Mvt.MODEPAIE = '' then
            begin
                LaListeAdd(GoListe1(Mvt, 'C', ' "' + Mvt.MODEPAIE + '" ', 35, 'E'));
                Result := False;
                //OK:=False ;
            end
            else
            begin
                LaListeAdd(GoListe1(Mvt, 'C', ' (' + Mvt.MODEPAIE + ') ', 39, 'E'));
                Result := False;
                //OK:=False ;
            end;
        end;
    end;

    if Pointable and (CGen.Naturegene = 'BQE') then //et banque
    begin
        if OkPourPoint then
        begin
            if Mvt.ECHE = 'X' then
            begin
                if (Mvt.MODEPAIE <> '') and (TCModP.values.IndexOf(Mvt.MODEPAIE) > -1) then
                begin
                    OkModP := True;
                    if Mvt.NUMECHE < 1 then Oknumeche := True;
                end
                else OkModP := False;
            end
            else OkEche := False;
        end
        else
        begin
            LaListeAdd(GoListe1(Mvt, 'C', ' (' + Mvt.EtatLettrage + ') ', 91, 'E'));
            Result := False;
            Exit;
        end;

        if not okEche then
        begin { Msg E_ECHE Error }
            LaListeAdd(GoListe2(Mvt, 'C', '', QEcr.FindField('' + FE + '_ECHE'), 'E'));
            Result := False;
        end;

        if not Oknumeche then
        begin { Msg E_NUMECHE Error }
            LaListeAdd(GoListe2(Mvt, 'C', '', QEcr.FindField('' + FE + '_NUMECHE'), 'E'));
            Result := False;
            //OK:=False ;
        end;

        if not OkModP then { Msg E_MODEPAIE vide ou inconnue }
        begin
            if Mvt.MODEPAIE = '' then
            begin
                LaListeAdd(GoListe1(Mvt, 'C', ' "' + Mvt.MODEPAIE + '" ', 35, 'E'));
                Result := False;
            end
            else
            begin
                LaListeAdd(GoListe1(Mvt, 'C', ' (' + Mvt.MODEPAIE + ') ', 39, 'E'));
                Result := False;
            end;
        end;
    end;
    //VerifInfoEche:=Ok and OkModP ;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... : 12/10/2005
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.VerifInfoDevise : Boolean;
begin

/// Mess error plus d'une fois car meme piece/ligne mais numeche different !!!

    Result := True; // C 60

    if Mvt.DEVISE <> '' then
    begin
        if not Existe(1, Mvt.DEVISE) then { Verif si Devise ok }
        begin
            LaListeAdd(GoListe1(Mvt, 'C', ' (' + Mvt.DEVISE + ') ', 53, 'E'));
            Result := False;
        end
        else
          // YMO 12/10/2005 FQ15422 Verifier que le journal est multidevise si ecriture en devise
          If Not EXISTESQL('SELECT * FROM JOURNAL WHERE J_JOURNAL = "'+Mvt.Journal+'" AND J_MULTIDEVISE = "X"')
          And (Mvt.DEVISE <> V_PGI.DevisePivot) then
          begin
            LaListeAdd(GoListe1(Mvt, 'C', ' (' + Mvt.DEVISE + ') ', 117, 'E'));
            Result := False;
          end;
    end
    else
    begin
        LaListeAdd(GoListe1(Mvt, 'C', ' "' + Mvt.DEVISE + '" ', 52, 'E'));
        Result := False;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.VerifInfoPieces : Boolean;
begin
    Result := True; // C 80

    if Mvt.NATUREPIECE <> '' then
    begin
        if not Existe(3, Mvt.NATUREPIECE) then
        begin
            LaListeAdd(GoListe1(Mvt, 'C', ' (' + Mvt.NATUREPIECE + ') ', 49, 'E'));
            Result := False;
        end
        else
        begin
            if not NATUREPIECECOMPTEOK(Mvt.NATUREPIECE, Mvt.GENERAL) then
            begin
                LaListeAdd(GoListe1(Mvt, 'C', ' (' + Mvt.NATUREPIECE + ') ' + ' (' + Mvt.GENERAL + ') ', 31, 'E'));
                Result := False;
            end;
            if Mvt.NATUREPIECE = 'ECC' then {Verif Montant en Ecart de change }
            begin
                if (Mvt.DEBIT = 0) and (Mvt.CREDIT = 0) and (Mvt.DEBITDEV = 0) and (Mvt.CREDITDEV = 0) then
                begin
                    LaListeAdd(GoListe1(Mvt, 'C', ' (' + Mvt.NATUREPIECE + ') ' + MsgLibel.Mess[37] + FloatToStr(Mvt.DEBIT) + ' ; ' + MsgLibel.Mess[38] + FloatToStr(Mvt.CREDIT), 41, 'E'));
                    Result := False;

                end;
                if (Mvt.DEBITDEV <> 0) and (Mvt.CREDITDEV <> 0) then
                begin
                    LaListeAdd(GoListe1(Mvt, 'C', ' (' + Mvt.NATUREPIECE + ') ' + MsgLibel.Mess[37] + FloatToStr(Mvt.DEBITDEV) + ' ; ' + MsgLibel.Mess[38] + FloatToStr(Mvt.CREDITDEV), 42, 'E'));
                    Result := False;
                end;

                if Mvt.Devise = V_PGI.DevisePivot then
                begin
                    LaListeAdd(GoListe1(Mvt, 'C', ' (' + Mvt.NATUREPIECE + ') ' + ' (' + Mvt.DEVISE + ') ' + MsgLibel.Mess[37] + FloatToStr(Mvt.DEBIT) + ' ; ' + MsgLibel.Mess[38] + FloatToStr(Mvt.CREDIT), 62, 'E'));
                    Result := False;
                end;
            end;
        end;
    end
    else
    begin
        LaListeAdd(GoListe1(Mvt, 'C', '', 48, 'E'));
        Result := False;
    end;

    if Mvt.NUMEROPIECE <= 0 then { Verif si NumPiece ok }
    begin
        if PourImport then LaListeAdd(GoListe2(Mvt, 'C', '', QEcr.FindField('IE_NUMPIECE'), 'E'))
        else LaListeAdd(GoListe2(Mvt, 'C', '', QEcr.FindField('E_NUMEROPIECE'), 'E'));
        Result := False;
    end;

    // FQ 13242
    if PourImport then { CA - 16/06/2004 - Ajout condition PourImport pour ne pas lancer le traitement
                              en phase de vérif lancée depuis la compta. En effet dans ce cas, QUALIFPIECE n'étant
                              pas renseigné, une erreur ressortait systématiquement }
    begin
        if not ((Mvt.QualifPiece = '') or (Mvt.QualifPiece = 'N') or (Mvt.QualifPiece = 'R') or (Mvt.QualifPiece = 'S') or (Mvt.QualifPiece = 'U') or (Mvt.QualifPiece = 'P') or (Mvt.QualifPiece = 'I')) then
        begin
            if PourImport then LaListeAdd(GoListe2(Mvt, 'C', '', QEcr.FindField('IE_QUALIFPIECE'), 'E'))
            else LaListeAdd(GoListe2(Mvt, 'C', '', QEcr.FindField('E_QUALIFPIECE'), 'E'));
            Result := False;
        end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.VerifTauxDev(MvtEcr : Boolean) : Boolean;
begin
    Result := True;

    if Mvt.DEVISE = '' then Exit;

    if Mvt.Devise = V_PGI.DevisePivot then
    begin
        if Mvt.TAUXDEV <> 1 then
        begin
            if MvtEcr then
                LaListeAdd(GoListe1(Mvt, 'C', ' (' + FloatTostr(Mvt.TAUXDEV) + ') ' + MsgLibel.Mess[37] + FloatToStr(Mvt.DEBIT) + ' ; ' + MsgLibel.Mess[38] + FloatToStr(Mvt.CREDIT), 77, 'E'))
            else
                LaListeAdd(GoListe1(Mvt, 'C', ' (' + FloatTostr(Mvt.TAUXDEV) + ') ' + MsgLibel.Mess[37] + FloatToStr(Mvt.DEBIT) + ' ; ' + MsgLibel.Mess[38] + FloatToStr(Mvt.CREDIT), 79, 'Y'));
            Result := False;
        end;
    end
    else
    begin
        if Mvt.NATUREPIECE = 'ECC' then exit;
        if (Mvt.DebitDev = 0) and (Mvt.CreditDev = 0) then Exit;
        (* GP le 25/05/98 : N'a aucun sens !!!
        Quot := DonneQuotite(Mvt.Devise);
        iF Abs(Mvt.Debit)>Abs(Mvt.Credit) then
        begin
            if Arrondi(Mvt.TAUXDEV,4)<>Arrondi((Quot*(Mvt.Debit/Mvt.DebitDev)),4) then
            begin
                if MvtEcr then LaListeAdd(GoListe1(Mvt,'C',' ('+FloatTostr(Mvt.TAUXDEV)+') '+' ('+Mvt.DEVISE+') '+MsgLibel.Mess[37]+FloatToStr(Mvt.DEBIT)+' ; '+MsgLibel.Mess[38]+FloatToStr(Mvt.CREDIT),78,'E'))
                else LaListeAdd(GoListe1(Mvt,'C',' ('+FloatTostr(Mvt.TAUXDEV)+') '+' ('+Mvt.DEVISE+') '+MsgLibel.Mess[37]+FloatToStr(Mvt.DEBIT)+' ; '+MsgLibel.Mess[38]+FloatToStr(Mvt.CREDIT),80,'Y'));
                result := False ;
            end;
        end
        else iF Abs(Mvt.Credit)>Abs(Mvt.Debit) then
        begin
            if Arrondi(Mvt.TAUXDEV,4)<>Arrondi((Quot*(Mvt.Credit/Mvt.CreditDev)),4) then
            begin
                if MvtEcr then LaListeAdd(GoListe1(Mvt,'C',' ('+FloatTostr(Mvt.TAUXDEV)+') '+' ('+Mvt.DEVISE+') '+MsgLibel.Mess[37]+FloatToStr(Mvt.DEBIT)+' ; '+MsgLibel.Mess[38]+FloatToStr(Mvt.CREDIT),78,'E'))
                else LaListeAdd(GoListe1(Mvt,'C',' ('+FloatTostr(Mvt.TAUXDEV)+') '+' ('+Mvt.DEVISE+') '+MsgLibel.Mess[37]+FloatToStr(Mvt.DEBIT)+' ; '+MsgLibel.Mess[38]+FloatToStr(Mvt.CREDIT),80,'Y'));
                result := False;
            end;
        end;
        *)
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.VerifEtab : Boolean;
begin
    Result := True; // xdoc

    if Mvt.Etablissement = '' then
    begin
        LaListeAdd(GoListe1(Mvt, 'C', '""', 69, 'E'));
        Result := False;
    end
    else
        if not (TCEtab.values.IndexOf(Mvt.Etablissement) > -1) then
    begin
        LaListeAdd(GoListe1(Mvt, 'C', ' (' + Mvt.Etablissement + ') ', 70, 'E'));
        Result := False;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.EcrAnoOk(Journal, Ecran : string) : Boolean;
begin
    Result := True;

    if (Journal <> 'ANO') and (Journal <> 'CLO') and (Journal <> 'ANA') then
    begin
        if Ecran <> 'N' then Result := False;
    end
    else
        if Journal = 'ANO' then
    begin
        if (Ecran <> 'H') and (ECRAN <> 'OAN') then Result := False;
    end
    else
        if Journal = 'ANA' then
    begin
        if (Ecran <> 'H') and (ECRAN <> 'OAN') then Result := False;
    end
    else
        if Journal = 'CLO' then
    begin
        if Ecran <> 'C' then Result := False;
    end;
end;

//********** Fonction de VERIFICATION analytique

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.VerifInfoAnal(Ventilable : Boolean) : Boolean;
var
    Ok : Boolean;
begin
    Ok := True;

    if Ventilable then { Gene ventil ? } // C 50
    begin
        if Mvt.ANA = 'X' then
        begin
//            if Mvt.NUMECHE=0 then Ok:=True Else Ok:=False ;
        end
        else ok := False;

        { E_ANA Error }
        if not Ok then LaListeAdd(GoListe2(Mvt, 'C', '', QEcr.FindField('' + FE + '_ANA'), 'E'));
    end;

    VerifInfoAnal := Ok;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.VerifCptAnal : Boolean;
var
    OkAx : Boolean;
begin
    Result := True;
    OkAx := True;

    if Mvt.Axe = '' then // D 11
    begin
        //if MAJ then Corrige()
        LaListeAdd(GoListe1(Mvt, 'C', '', 60, 'Y'));
        Result := False;
        Exit;
    end
    else
        if (Char(Mvt.Axe[2]) <> '1') and (Char(Mvt.Axe[2]) <> '2') and (Char(Mvt.Axe[2]) <> '3') and (Char(Mvt.Axe[2]) <> '4') and (Char(Mvt.Axe[2]) <> '5') then
    begin
        LaListeAdd(GoListe1(Mvt, 'C', '(' + Mvt.Axe + ')', 61, 'Y'));
        Result := False;
        Exit;
    end;

    if Mvt.Section = '' then // D 12
    begin
        LaListeAdd(GoListe1(Mvt, 'C', '', 59, 'Y'));
        Result := False;
        Exit;
    end;

    { Y a t-il un code section existant sur un axe ? }
    if not ExisteSec(Mvt.Axe, Mvt.Section) then
    begin
        { Section n'existe pas, Affichage dans la lise }
        LaListeAdd(GoListe1(Mvt, 'C', ' (' + Mvt.Section + ' : ' + Mvt.Axe + ') ', 11, 'Y'));
        Result := False;
    end;

    if Mvt.TYPEANALYTIQUE = 'X' then
        if Mvt.GENERAL = '' then Exit; {En Anal Pur le gene peut ne pas etre renseigné}

    if Mvt.General = '' then // D 13
    begin
        LaListeAdd(GoListe1(Mvt, 'C', ' "' + Mvt.General + '" ', 58, 'Y'));
        Result := False;
        Exit;
    end;

    Fillchar(CGen, SizeOf(CGen), #0);
    InfosCptGen(Mvt.General);
    if not Cgen.Ventilable[StrToInt(Char(Mvt.Axe[2]))] then OkAx := False;

    { Verif Gene ventil  }
    if not OkAx then
    begin
        LaListeAdd(GoListe1(Mvt, 'C', ' (' + Mvt.GENERAL + ') (' + Mvt.Axe + ')', 28, 'Y'));
        Result := False;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.VerifInfoJalAnal : Boolean;
begin
    Result := True;
    Fillchar(InfAutres, SizeOf(InfAutres), #0);

    if Mvt.Journal = '' then // D 21
    begin
        LaListeAdd(GoListe1(Mvt, 'C', ' "' + Mvt.JOURNAL + '" ', 47, 'Y'));
        Result := False;
        Exit;
    end
    else
    begin
        if not Existe(2, Mvt.JOURNAL) then { Verif si Journal ok } // D 22
        begin
            LaListeAdd(GoListe1(Mvt, 'C', ' "' + Mvt.JOURNAL + '" ', 36, 'Y'));
            Result := False;
            Exit;
        end;
    end;

    if Mvt.TYPEANALYTIQUE = 'X' then { Journal D' OD Analytique }
    begin
        if (InfAutres.J_NATUREJAL <> 'ODA') and (InfAutres.J_NATUREJAL <> 'ANA') then // D 23
        begin
            { Journal doit être de nature ODA et ODA}
            LaListeAdd(GoListe1(Mvt, 'C', '', 26, 'Y'));
            Result := False;
        end
        else
        begin
            if InfAutres.J_Axe <> Mvt.AXE then // D 24
            begin
                { Journal n'ayant pas le meme Axe ou inexistant }
                LaListeAdd(GoListe1(Mvt, 'C', ' (' + Mvt.AXE + ') ', 27, 'Y'));
                Result := False;
            end;
        end;
    end
    else { Analytique Comptable }
    begin
        if (InfAutres.J_NATUREJAL = 'ODA') or (InfAutres.J_NATUREJAL = 'ANA') then // D 25
        begin
            { Journal est de nature ODA ou ANA }
            LaListeAdd(GoListe1(Mvt, 'C', '', 40, 'Y'));
            Result := False;
        end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.VerifInfoExoAnal : Boolean;
begin
    Result := True;

    if PourImport then
    begin
        InfAutres.EX_DATEEXO[0] := VH^.EnCours.Deb;
        InfAutres.EX_DATEEXO[1] := VH^.EnCours.Fin;
        if (Mvt.DATECOMPTABLE >= VH^.Precedent.Deb) and (Mvt.DATECOMPTABLE <= VH^.Precedent.Fin) and (VH^.Precedent.Code <> '') then
        begin
            InfAutres.EX_DATEEXO[0] := VH^.Precedent.Deb;
            InfAutres.EX_DATEEXO[1] := VH^.Precedent.Fin;
        end
        else
            if (Mvt.DATECOMPTABLE >= VH^.EnCours.Deb) and (Mvt.DATECOMPTABLE <= VH^.EnCours.Fin) then
        begin
            InfAutres.EX_DATEEXO[0] := VH^.EnCours.Deb;
            InfAutres.EX_DATEEXO[1] := VH^.EnCours.Fin;
        end
        else
            if (Mvt.DATECOMPTABLE >= VH^.Suivant.Deb) and (Mvt.DATECOMPTABLE <= VH^.Suivant.Fin) and (VH^.Suivant.Code <> '') then
        begin
            InfAutres.EX_DATEEXO[0] := VH^.Suivant.Deb;
            InfAutres.EX_DATEEXO[1] := VH^.Suivant.Fin;
        end;

        Exit;
    end;

    if not OkExo(Mvt.Exercice, Mvt.DateComptable) then // D 50
    begin
        LaListeAdd(GoListe2(Mvt, 'C', '', QAnal.FindField('Y_EXERCICE'), 'Y'));
        Result := False;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.VerifInfoEcrAnoAnal : Boolean;
begin
    Result := True;

    if not EcrAnoOk(InfAutres.J_NATUREJAL, Mvt.ECRANOUVEAU) then // D 30
    begin
        LaListeAdd(GoListe2(Mvt, 'C', '( ' + Mvt.ECRANOUVEAU + ' )', QAnal.FindField('' + FY + '_ECRANOUVEAU'), 'Y'));
        Result := False;
    end;

    if InfAutres.J_NATUREJAL = 'ANO' then
    begin
        if PourImport then Exit;

        if Mvt.DateComptable <> InfAutres.EX_DATEEXO[0] then // xdoc
        begin
            LaListeAdd(GoListe1(Mvt, 'C', '(' + DateToStr(InfAutres.EX_DATEEXO[0]) + ' )', 68, 'Y'));
            Result := False;
        end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.VerifInfoPieceAnal : Boolean;
begin
    Result := True; // D 40

    if Mvt.TYPEANALYTIQUE = 'X' then { Anal Pur }
    begin
        if Mvt.NUMLIGNE <> 0 then
        begin
            { Numero de ligne ne doit pas etre >0, Affichage dans la lise }
            LaListeAdd(GoListe2(Mvt, 'C', '', QAnal.FindField('' + FY + '_NUMLIGNE'), 'Y'));
            Result := False;
        end;
    end;

    if Mvt.NUMVENTIL < 1 then
    begin
        LaListeAdd(GoListe2(Mvt, 'C', '', QAnal.FindField('' + FY + '_NUMVENTIL'), 'Y'));
        Result := False;
    end;

    if Mvt.NATUREPIECE <> '' then
    begin
        if not Existe(3, Mvt.NATUREPIECE) then
        begin
            LaListeAdd(GoListe1(Mvt, 'C', ' (' + Mvt.NATUREPIECE + ') ', 51, 'Y'));
            Result := False;
        end
        else
        begin (*Prevoir aussi Proc. NaturePieceJal si Ok*)
            if not NATUREPIECECOMPTEOK(Mvt.NATUREPIECE, Mvt.GENERAL) then
            begin
                LaListeAdd(GoListe1(Mvt, 'C', ' (' + Mvt.NATUREPIECE + ') ' + ' (' + Mvt.GENERAL + ') ', 32, 'Y'));
                Result := False;
            end;

            if Mvt.NATUREPIECE = 'ECC' then {Verif Montant en Ecart de change }
            begin
                if (Mvt.DEBIT = 0) and (Mvt.CREDIT = 0) and (Mvt.DEBITDEV = 0) and (Mvt.CREDITDEV = 0) then
                begin
                    LaListeAdd(GoListe1(Mvt, 'C', ' (' + Mvt.NATUREPIECE + ') ', 43, 'Y'));
                    Result := False;
                end;

                if (Mvt.DEBITDEV <> 0) and (Mvt.CREDITDEV <> 0) then
                begin
                    LaListeAdd(GoListe1(Mvt, 'C', ' (' + Mvt.NATUREPIECE + ') ', 44, 'Y'));
                    Result := False;
                end;

                if Mvt.Devise = V_PGI.DevisePivot then
                begin
                    LaListe.Add(GoListe1(Mvt, 'C', ' (' + Mvt.NATUREPIECE + ') ' + ' (' + Mvt.DEVISE + ') ', 63, 'Y'));
                    Result := False;
                end;
            end;
        end;
    end
    else
    begin
        LaListeAdd(GoListe1(Mvt, 'C', ' "' + Mvt.NATUREPIECE + '" ', 50, 'Y'));
        Result := False;
    end;

    if Mvt.NUMEROPIECE <= 0 then
    begin
        if PourImport then LaListeAdd(GoListe2(Mvt, 'C', '', QAnal.FindField('IE_NUMPIECE'), 'Y'))
        else LaListeAdd(GoListe2(Mvt, 'C', '', QAnal.FindField('Y_NUMEROPIECE'), 'Y'));

        Result := False;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... : 12/10/2005
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.VerifInfoDeviseAnal : Boolean;
begin
    Result := True; // D 60

    if Mvt.DEVISE <> '' then
    begin
        if not Existe(1, Mvt.DEVISE) then { Verif si Devise ok }
        begin
            LaListeAdd(GoListe1(Mvt, 'C', ' (' + Mvt.EXERCICE + ') ', 55, 'Y'));
            Result := False;
        end
        else
          // YMO 12/10/2005 FQ15422 Verifier que le journal est multidevise si ecriture en devise
          If Not EXISTESQL('SELECT * FROM JOURNAL WHERE J_JOURNAL = "'+Mvt.Journal+'" AND J_MULTIDEVISE = "X"')
          And (Mvt.DEVISE <> V_PGI.DevisePivot) then
          begin
            LaListeAdd(GoListe1(Mvt, 'C', ' (' + Mvt.DEVISE + ') ', 118, 'Y'));
            Result := False;
          end;
    end
    else
    begin
        LaListeAdd(GoListe1(Mvt, 'C', ' "' + Mvt.DEVISE + '" ', 54, 'Y'));
        Result := False;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.VerifEtabAnal : Boolean;
begin
    Result := True; // xdoc

    if Mvt.Etablissement = '' then
    begin
        LaListeAdd(GoListe1(Mvt, 'C', '""', 71, 'Y'));
        Result := False;
    end
    else
        if not (TCEtab.values.IndexOf(Mvt.Etablissement) > -1) then
    begin
        LaListeAdd(GoListe1(Mvt, 'C', ' (' + Mvt.Etablissement + ') ', 72, 'Y'));
        Result := False;
    end;
end;

//********** Fonction de VERIFICATION du lettrage

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.VerifMvtNoLet : Boolean; { Appelé par SQLLet }
var
    OkLet : Boolean;
begin
    Result := True; // E 10

    if (QLettrage.Fields[2].AsString <> '') then Exit; // E_Lettrage

    { E_LETTRAGE and (E_ETATLETTRAGE or E_ETATLETTRAGE) }
    OkLet := (QLettrage.FieldS[2].AsString = '') and ((QLettrage.FindField('E_ETATLETTRAGE').AsString = 'AL') or (QLettrage.FindField('E_ETATLETTRAGE').AsString = 'RI'));

    if not OkLet then
    begin
        LaListeAdd(GoListe1(Mvt, 'L', MsgLibel.Mess[15] + FloatToStr(QLettrage.FindField('E_DEBIT').AsFloat) + ' ; ' + MsgLibel.Mess[16] + FloatToStr(QLettrage.FindField('E_CREDIT').AsFloat), 22, 'E'));
        Result := False;
        exit;
    end
    else
    begin
        if Mvt.LETTRAGEDEV <> 'X' then
        begin
            if (Mvt.COUVERTURE <> 0) then
            begin
                LaListeAdd(GoListe2(Mvt, 'L', ' (' + FloatToStr(Mvt.COUVERTURE) + ') ' + '"' + Mvt.DEVISE + '" ' + MsgLibel.Mess[37] + FloatToStr(Mvt.DEBIT) + ' ; ' + MsgLibel.Mess[38] + FloatToStr(Mvt.CREDIT), QLettrage.FindField('' + FE + '_COUVERTURE'), 'E'));
                Result := False;
            end;
        end
        else
            if Mvt.LETTRAGEDEV = 'X' then
        begin
            if (Mvt.COUVERTUREDEV <> 0) then
            begin
                LaListeAdd(GoListe2(Mvt, 'L', ' (' + FloatToStr(Mvt.COUVERTUREDEV) + ') ' + '"' + Mvt.DEVISE + '" ' + MsgLibel.Mess[37] + FloatToStr(Mvt.DEBIT) + ' ; ' + MsgLibel.Mess[38] + FloatToStr(Mvt.CREDIT), QLettrage.FindField('' + FE + '_COUVERTUREDEV'), 'E'));
                Result := False;
            end;
        end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.VerifMvtLet : Boolean; { Appelé par SQLLet }
var
    OkLet, OkDatePaquet : Boolean;
    DecDev : Byte;
begin
    Result := True; // E 20

    if (QLettrage.FindField('' + FE + '_LETTRAGE').AsString = '') then Exit;

    OkLet := (Majuscule(QLettrage.FindField('' + FE + '_LETTRAGE').AsString) and (QLettrage.FindField('' + FE + '_ETATLETTRAGE').AsString = 'TL')) or (not Majuscule(QLettrage.FindField('' + FE + '_LETTRAGE').AsString) and (QLettrage.FindField('' + FE + '_ETATLETTRAGE').AsString = 'PL'));

    if not OkLet then
    begin
        LaListeAdd(GoListe1(Mvt, 'L', MsgLibel.Mess[37] + FloatToStr(QLettrage.FindField('' + FE + '_DEBIT').AsFloat) + ' ; '
            + MsgLibel.Mess[38] + FloatToStr(QLettrage.FindField('' + FE + '_CREDIT').AsFloat), 23, 'E'));
        Result := False;
        Exit;
    end;

    OkDatePaquet := (QLettrage.FindField('' + FE + '_DATECOMPTABLE').AsDateTime >= QLettrage.FindField('' + FE + '_DATEPAQUETMIN').AsDateTime) and (QLettrage.FindField('' + FE + '_DATECOMPTABLE').AsDateTime <= QLettrage.FindField('' + FE + '_DATEPAQUETMAX').AsDateTime);

    if not OkDatePaquet then
    begin
        LaListeAdd(GoListe1(Mvt, 'L', MsgLibel.Mess[37] + FloatToStr(QLettrage.FindField('' + FE + '_DEBIT').AsFloat) + ' ; '
            + MsgLibel.Mess[38] + FloatToStr(QLettrage.FindField('' + FE + '_CREDIT').AsFloat), 73, 'E'));
        Result := False;
    end;

    if Mvt.ETATLETTRAGE = 'PL' then
    begin
        {--- en Devise Pivot ---}
        if Mvt.LETTRAGEDEV <> 'X' then
        begin
            if Mvt.DEBIT <> 0 then
            begin
                if Mvt.COUVERTURE < 0 then
                    if Mvt.COUVERTURE > Mvt.DEBIT then
                    begin
                        LaListeAdd(GoListe2(Mvt, 'L', ' (' + FloatToStr(Mvt.COUVERTURE) + ') ' + '"' + Mvt.DEVISE + '" ' + MsgLibel.Mess[37] + FloatToStr(Mvt.DEBIT) + ' ; ' + MsgLibel.Mess[38] + FloatToStr(Mvt.CREDIT), QLettrage.FindField('' + FE + '_COUVERTURE'), 'E'));
                        Result := False;
                    end;
            end;

            if Mvt.CREDIT <> 0 then
            begin
                if Mvt.COUVERTURE < 0 then
                    if Mvt.COUVERTURE > Mvt.CREDIT then
                    begin
                        LaListeAdd(GoListe2(Mvt, 'L', ' (' + FloatToStr(Mvt.COUVERTURE) + ') ' + '"' + Mvt.DEVISE + '" ' + MsgLibel.Mess[37] + FloatToStr(Mvt.DEBIT) + ' ; ' + MsgLibel.Mess[38] + FloatToStr(Mvt.CREDIT), QLettrage.FindField('' + FE + '_COUVERTURE'), 'E'));
                        Result := False;
                    end;
            end;
        end
        else
            if Mvt.LETTRAGEDEV = 'X' then
        begin
            {--- en Devise  ---}
            if Mvt.DEBITDEV <> 0 then
            begin
                if Mvt.COUVERTUREDEV < 0 then
                    if Mvt.COUVERTUREDEV > Mvt.DEBITDEV then
                    begin
                        LaListeAdd(GoListe2(Mvt, 'L', ' (' + FloatToStr(Mvt.COUVERTUREDEV) + ') ' + '"' + Mvt.DEVISE + '" ' + MsgLibel.Mess[74] + FloatToStr(Mvt.DEBITDEV) + ' ; ' + MsgLibel.Mess[75] + FloatToStr(Mvt.CREDITDEV), QLettrage.FindField('' + FE + '_COUVERTUREDEV'), 'E'));
                        Result := False;
                    end;
            end;

            if Mvt.CREDITDEV <> 0 then
            begin
                if Mvt.COUVERTUREDEV < 0 then
                    if Mvt.COUVERTUREDEV > Mvt.CREDITDEV then
                    begin
                        LaListeAdd(GoListe2(Mvt, 'L', ' (' + FloatToStr(Mvt.COUVERTUREDEV) + ') ' + '"' + Mvt.DEVISE + '" ' + MsgLibel.Mess[74] + FloatToStr(Mvt.DEBITDEV) + ' ; ' + MsgLibel.Mess[75] + FloatToStr(Mvt.CREDITDEV), QLettrage.FindField('' + FE + '_COUVERTUREDEV'), 'E'));
                        Result := False;
                    end;
            end;
        end;
    end
    else
        if Mvt.ETATLETTRAGE = 'TL' then
    begin
        {--- en Devise Pivot ---}
        if Mvt.LETTRAGEDEV <> 'X' then
        begin
            if Mvt.DEBIT <> 0 then
            begin
                if Arrondi(Mvt.COUVERTURE, V_PGI.OkdecV) <> Arrondi(Abs(Mvt.DEBIT), V_PGI.OkdecV) then // FFF 181
                begin
                    LaListeAdd(GoListe2(Mvt, 'L', ' (' + FloatToStr(Mvt.COUVERTURE) + ') ' + '"' + Mvt.DEVISE + '" ' + MsgLibel.Mess[37] + FloatToStr(Mvt.DEBIT) + ' ; ' + MsgLibel.Mess[38] + FloatToStr(Mvt.CREDIT), QLettrage.FindField('' + FE + '_COUVERTURE'), 'E'));
                    Result := False;
                end;
            end;

            if Mvt.CREDIT <> 0 then
            begin
                if Arrondi(Mvt.COUVERTURE, V_PGI.OkdecV) <> Arrondi(Abs(Mvt.CREDIT), V_PGI.OkdecV) then // FFF 181
                begin
                    LaListeAdd(GoListe2(Mvt, 'L', ' (' + FloatToStr(Mvt.COUVERTURE) + ') ' + '"' + Mvt.DEVISE + '" ' + MsgLibel.Mess[37] + ' ' + FloatToStr(Mvt.DEBIT) + ' ; ' + MsgLibel.Mess[38] + ' ' + FloatToStr(Mvt.CREDIT), QLettrage.FindField('' + FE + '_COUVERTURE'), 'E'));
                    Result := False;
                end;
            end;
        end;
        {--- en Devise  ---}
        if Mvt.LETTRAGEDEV = 'X' then
        begin
            if Mvt.Devise <> V_PGI.DevisePivot then
            begin
                TrouveDecDev(Mvt.Devise, DecDev);
            end
            else DecDev := V_PGI.OkdecV;

            if Mvt.DEBITDEV <> 0 then
            begin
                if Arrondi(Mvt.COUVERTUREDEV, DecDev) <> Arrondi(Abs(Mvt.DEBITDEV), DecDev) then // 14747
                begin
                    LaListeAdd(GoListe2(Mvt, 'L', ' (' + FloatToStr(Mvt.COUVERTUREDEV) + ') ' + '"' + Mvt.DEVISE + '" ' + MsgLibel.Mess[74] + ' ' + FloatToStr(Mvt.DEBITDEV) + ' ; ' + MsgLibel.Mess[75] + ' ' + FloatToStr(Mvt.CREDITDEV), QLettrage.FindField('' + FE + '_COUVERTUREDEV'), 'E'));
                    Result := False;
                end;
            end;

            if Mvt.CREDITDEV <> 0 then
            begin
                if Arrondi(Mvt.COUVERTUREDEV, DecDev) <> Arrondi(Abs(Mvt.CREDITDEV), DecDev) then // 14747
                begin
                    LaListeAdd(GoListe2(Mvt, 'L', ' (' + FloatToStr(Mvt.COUVERTUREDEV) + ') ' + '"' + Mvt.DEVISE + '" ' + MsgLibel.Mess[74] + ' ' + FloatToStr(Mvt.DEBITDEV) + ' ; ' + MsgLibel.Mess[75] + ' ' + FloatToStr(Mvt.CREDITDEV), QLettrage.FindField('' + FE + '_COUVERTUREDEV'), 'E'));
                    Result := False;
                end;
            end;
        end;
    end;

end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.VerifLetDev : Boolean; { Appelé par SQLVerifLet(1) }
var
    Q : TQuery;
    OkDev : Boolean;
    ErLetDev : TErrLetDev;
    St : string;
    LaDevise : String3;
begin
// E 30
    OkDev := True;

    st := ' Select ' + FE + '_LETTRAGEDEV, ' + FE + '_DEVISE ';
    St := St + ' From ' + FICECR + ' '
        + ' where ' + FE + '_QUALIFPIECE="N" '
        + ' AND ' + FE + '_AUXILIAIRE="' + QLetDev.FindField('' + FE + '_AUXILIAIRE').AsString + '" '
        + ' AND ' + FE + '_GENERAL="' + QLetDev.FindField('' + FE + '_GENERAL').AsString + '" '
        + ' AND ' + FE + '_LETTRAGE="' + QLetDev.findField('' + FE + '_LETTRAGE').AsString + '" '
        + ' And ' + FE + '_eche="X" and ' + FE + '_etatlettrage<>"" and ' + FE + '_etatlettrage<>"RI" '
        + ' And ' + FE + '_Ecranouveau<>"OAN" and ' + FE + '_ecranouveau<>"C" ';
    if VH^.ExoV8.Code <> '' then St := St + ' And ' + FE + '_DATECOMPTABLE>="' + UsDateTime(VH^.ExoV8.Deb) + '" ';
    St := St + ' group BY ' + FE + '_AUXILIAIRE, ' + FE + '_GENERAL, ' + FE + '_ETATLETTRAGE, ' + FE + '_LETTRAGE,' + FE + '_LETTRAGEDEV, ' + FE + '_DEVISE ';

    Q := OpenSql(St, True);

    LaDevise := '';

    while not Q.Eof do
    begin
        if Q.FieldS[0].AsString = 'X' then // E_Lettragedev
            if LaDevise <> '' then
                if Q.Fields[1].AsString <> LaDevise then //E_Devise
                begin
                    OkDev := False;
                    Break;
                end;
        LaDevise := Q.Fields[1].AsString;
        Q.Next; //E_Devise
    end;

    if not OkDev then
    begin { E_LETTRAGEDEV Error }
        if Mvt <> nil then
        begin
            Mvt.FREE;
            Mvt := nil;
        end;
        EnregLet(QLetDev);
        ErLetDev := TErrLetDev.Create;
        LaListeAdd(GoListe2(Mvt, 'L', '', QLetDev.FindField('' + FE + '_LETTRAGEDEV'), 'L'));
        ErLetDev.Lettres := QLetDev.findfield('' + FE + '_LETTRAGE').AsString;
        ErLetDev.Auxi := QLetDev.FindField('' + FE + '_AUXILIAIRE').AsString;
        ErLetDev.Gene := QLetDev.findField('' + FE + '_GENERAL').AsString;
        ErLetDev.EtatLet := QLetDev.FindField('' + FE + '_ETATLETTRAGE').AsString;
        LetListe.AddObject('', ErLetDev);
    end;
    Ferme(Q);
    Result := OkDev;
end;

//********** Fonction de VERIFICATION budgetaire

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.VerifCompteBud : Boolean;
begin
    Result := True;

    { Verif si Cpt Budgétaire est renseigné ou existe }
    if MvtBud.BUDGENE = '' then // C 11
    begin
        LaListeAdd(GoListe1(MvtBud, 'B', '', 102, 'B'));
        Result := False;
        Exit;
    end;

    { Y a t-il un code section existant sur un axe ? }
    if not ExisteCpteBud(MvtBud.BUDGENE) then
    begin
        { Section n'existe pas, Affichage dans la lise }
        LaListeAdd(GoListe1(MvtBud, 'B', ' "' + MvtBud.BUDGENE + '" ', 103, 'B'));
        Result := False;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.VerifInfoJalBud : Boolean;
begin
    Result := True;

    { Verif si journal existe }
    if MvtBud.BudJal = '' then // C 11
    begin
        LaListeAdd(GoListe1(MvtBud, 'B', '', 100, 'B'));
        Result := False;
        Exit;
    end;

    { Y a t-il un code section existant sur un axe ? }
    if not ExisteJalBud(MvtBud.BUDJAL) then
    begin
        { Section n'existe pas, Affichage dans la lise }
        LaListeAdd(GoListe1(MvtBud, 'B', ' "' + MvtBud.BudJal + '" ', 101, 'B'));
        Result := False;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.VerifInfoExoBud : Boolean;
begin
    Result := True; //E_EXERCICE Error   // C90

    if PourImport then
    begin
        InfAutres.EX_DATEEXO[0] := VH^.EnCours.Deb;
        InfAutres.EX_DATEEXO[1] := VH^.EnCours.Fin;

        if (MvtBud.DATECOMPTABLE >= VH^.Precedent.Deb) and (MvtBud.DATECOMPTABLE <= VH^.Precedent.Fin) and (VH^.Precedent.Code <> '') then
        begin
            InfAutres.EX_DATEEXO[0] := VH^.Precedent.Deb;
            InfAutres.EX_DATEEXO[1] := VH^.Precedent.Fin;
        end
        else
            if (MvtBud.DATECOMPTABLE >= VH^.EnCours.Deb) and (MvtBud.DATECOMPTABLE <= VH^.EnCours.Fin) then
        begin
            InfAutres.EX_DATEEXO[0] := VH^.EnCours.Deb;
            InfAutres.EX_DATEEXO[1] := VH^.EnCours.Fin;
        end
        else
            if (MvtBud.DATECOMPTABLE >= VH^.Suivant.Deb) and (MvtBud.DATECOMPTABLE <= VH^.Suivant.Fin) and (VH^.Suivant.Code <> '') then
        begin
            InfAutres.EX_DATEEXO[0] := VH^.Suivant.Deb;
            InfAutres.EX_DATEEXO[1] := VH^.Suivant.Fin;
        end;

        exit;
    end;

    if not OkExo(MvtBud.EXERCICE, MvtBud.DATECOMPTABLE) then
    begin
        LaListeAdd(GoListe1(MvtBud, 'B', ' (' + MvtBud.Exercice + ') ', 95, 'E'));
        Result := False;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.VerifInfoSecBud : Boolean;
begin
    Result := True;

    if MvtBud.Axe = '' then // D 11
    begin
        //if MAJ then Corrige()
        LaListeAdd(GoListe1(MvtBud, 'B', '', 98, 'B'));
        Result := False;
        Exit;
    end
    else
        if (Char(MvtBud.Axe[2]) <> '1') and (Char(MvtBud.Axe[2]) <> '2') and (Char(MvtBud.Axe[2]) <> '3') and (Char(MvtBud.Axe[2]) <> '4') and (Char(MvtBud.Axe[2]) <> '5') then
    begin
        LaListeAdd(GoListe1(MvtBud, 'B', ' "' + MvtBud.BudSect + '" (' + MvtBud.Axe + ')', 99, 'Y'));
        Result := False;
        Exit;
    end;

    if MvtBud.BUDSECT = '' then // D 12
    begin
        LaListeAdd(GoListe1(MvtBud, 'B', '', 97, 'B'));
        Result := False;
        Exit;
    end;

    { Y a t-il un code section existant sur un axe ? }
    if not ExisteSecBud(MvtBud.Axe, MvtBud.BUDSECT) then
    begin
        { Section n'existe pas, Affichage dans la lise }
        LaListeAdd(GoListe1(MvtBud, 'B', ' "' + MvtBud.BudSect + '" (' + MvtBud.Axe + ') ', 96, 'B'));
        Result := False;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TFVerCpta.VerifEtabBud : Boolean;
begin
    Result := True; // xdoc

    if MvtBud.Etablissement = '' then
    begin
        LaListeAdd(GoListe1(MvtBud, 'B', '""', 69, 'B'));
        Result := False;
    end
    else
        if not (TCEtab.values.IndexOf(MvtBud.Etablissement) > -1) then
    begin
        LaListeAdd(GoListe1(MvtBud, 'B', ' (' + MvtBud.Etablissement + ') ', 70, 'B'));
        Result := False;
    end;
end;

//********** Fonction non utilisé

(*
procedure TFVerCpta.MAJMVT(Champ, Valeur : string);
var
    St : string;
begin
    ST := ' Update ANALYTIQ set ' + Champ + '="' + Valeur + '"';
    St := St + ' where Y_JOURNAL="' + Mvt.Journal + '" and Y_DATECOMPTABLE="' + USDATETIME(Mvt.DATECOMPTABLE) + '" ';
    St := St + ' and Y_EXERCICE="' + Mvt.Exercice + '" and Y_NUMEROPIECE=' + IntToStr(Mvt.NUMEROPIECE) + ' ';
    St := St + ' and Y_NUMLIGNE=' + IntToStr(Mvt.NUMLIGNE) + ' and Y_DEVISE="' + Mvt.DEVISE + '" ';
    St := St + ' and Y_ETABLISSEMENT="' + Mvt.ETABLISSEMENT + '" and Y_AXE="' + Mvt.Axe + '" ';

    ExecuteSql(St);
end;

function TFVerCpta.VerifInfoConfAnal : Boolean;
begin
    Result := True; // D 70

    if PourImport then Exit;

    if Mvt.CONFIDENTIEL <> '0' then
        if Mvt.CONFIDENTIEL <> '1' then
        begin
            LaListeAdd(GoListe2(Mvt, 'C', '', QAnal.FindField('Y_CONFIDENTIEL'), 'Y'));
            Result := False;
        end;
end;

function TFVerCpta.DevBqe(Cpt : string) : string;
var
    i : integer;
    Inf : TEnregInfos;
begin
    Result := '';

    for i := 0 to ListeBQE.Count - 1 do
    begin
        if ListeBQE.Objects[i] <> nil then
        begin
            Inf := TEnregInfos(ListeBQE.Objects[i]);
            if (Inf.Inf1 = Cpt) then
            begin
                Result := Inf.Inf2;
                Break;
            end;
        end
        else Break;
    end;
end;

function TFVerCpta.VerifInfoConf : Boolean;
begin
    Result := True; // C95

    if Mvt.CONFIDENTIEL <> '0' then
    begin
        if Mvt.CONFIDENTIEL <> '1' then
        begin
            LaListeAdd(GoListe2(Mvt, 'C', ' (' + Mvt.CONFIDENTIEL + ') ', QEcr.FindField('' + FE + '_CONFIDENTIEL'), 'E'));
            Result := False;
        end;
    end;
end;

function TFVerCpta.DonneQuotite(Dev : string) : Double;
var
    i : Integer;
    Inf : TEnregInfos;
begin
    Result := 0;

    for i := 0 to ListeDEV.Count - 1 do
    begin
        if ListeDEV.Objects[i] <> nil then
        begin
            Inf := TEnregInfos(ListeDEV.Objects[i]);
            if (Inf.Inf1 = Dev) then
            begin
                Result := Inf.Mont;
                Break;
            end;
        end
        else Break;
    end;
end;

procedure TFVerCpta.GoListeImp(Chrono : Integer;Code : string);
var
    ImpError : TERRORIMP;
begin
    if ImpListe.Count <> 0 then
    begin
        ImpError := TERRORIMP(ImpListe.Objects[ImpListe.Count - 1]);
        if ImpError.NumChrono = Chrono then exit;
    end;

    ImpError := TERRORIMP.Create;
    ImpError.NumChrono := Chrono;
    ImpError.NumErreur := Code;
    ImpListe.AddObject('', ImpError);
end;
*)

//================================================================================
// Initialization
//================================================================================

end.
