unit CloPerio;

interface

uses
    Windows,
    Messages,
    SysUtils,
    Classes,
    Controls,
    Forms,
    Dialogs,
    StdCtrls,                                             
    Buttons,
    Hctrls,
    ComCtrls,
    ExtCtrls,
{$IFDEF EAGLCLIENT}
    UTOB,
{$ELSE}
    {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
    hmsgbox,
    Ent1,
    HQry,
    HEnt1,
    Mask,
    Hcompte,
    ParamDat,
   // Echeance, // ProchaineDate
    HTB97,
    HPanel,
    UiUtil,
    Paramsoc, // SetParamSoc, GetParamSocSecur
    {$IFDEF NETEXPERT}
    UNeACtions,
    {$ENDIF}
    {$IFDEF MODENT1}
    CPTypeCons,
    {$ENDIF MODENT1}
    HSysMenu;

procedure CloPer;
procedure AnnuleCloPer;
function CtrlTout(Exo : TExoDate) : Byte;
function CtrlTo(Exo : TExoDate) : TTraitement;
// ajout me pour clôture périodique dans le cas synchro
procedure CloPerSynchro(DateArr : string);

type
    TCloPer = class(TForm)
        MsgBleme : THMsgBox;
        Panel1 : TPanel;
        HLabel4 : THLabel;
        FExercice : THValComboBox;
        HLabel3 : THLabel;
        FDateCpta1 : THValComboBox;
        FDateCpta2 : THValComboBox;
        Resume : TMemo;
        Confirmation : THMsgBox;
        Morceaux : THMsgBox;
        HTexte1 : THLabel;
        Htexte2 : THLabel;
        HMTrad : THSystemMenu;
        HDernClo : THLabel;
        HPatience : THLabel;
        Dock971 : TDock97;
        HPB : TToolWindow97;
        BValider : TToolbarButton97;
        BFerme : TToolbarButton97;
        BAide : TToolbarButton97;
        BArchive: TCheckBox;
    FDateCptaExo: THValComboBox;
        procedure FormShow(Sender : TObject);
        procedure FExerciceChange(Sender : TObject);
        procedure BValiderClick(Sender : TObject);
        procedure BFermeClick(Sender : TObject);
        procedure FormClose(Sender : TObject;var Action : TCloseAction);
        procedure FormCreate(Sender : TObject);
        procedure BAideClick(Sender : TObject);
    procedure BArchiveClick(Sender: TObject);
    procedure FDateCpta1Change(Sender: TObject);
    private
    { Déclarations privées }
        FMois, Annee : Word; { Debut de mois et Année de L'exercice }
        DateExo : TExoDate;
        DateClotActu, DateClotNouv : TDateTime;
        NbLigne : Integer;
        CtrlExterne, CFait : Boolean;
        ResultatCtrlExterne : Byte;
        ExoExterne : TExoDate;
        PourAuditCloture : Boolean;
        gbCloture, CloParSynchro : Boolean;
{$IFDEF CERTIFNF}
        SessValidation : Integer;
        ChecksumValidation : String;
{$ENDIF}        
        function CtrlDate(DateArr : string = '') : Boolean;
        function MvtExotiquesOK(DAte1, Date2 : TDateTime) : Boolean;
        procedure LanceCloture;
        procedure LanceDecloture; // VL
        procedure RempliComboExo;
//        function DonneJDate(Periodes : THValComBoBox;DebOuFin : Boolean;Mo, An : word) : TDateTime;
        procedure GenPerExo(var DebMois, LAnnee : Word;Exo : string = '');
        function SQLDateClo : TDateTime;
        procedure RefreshMemo;
        procedure ClotureExoPrecedent (Exo :string='');
        procedure DeclotureExoSuivant;
        procedure DoDecloture;
        procedure DoCloture(Exo : string = '');
    public
    { Déclarations publiques }
    end;

implementation
{$R *.DFM}


uses
  {$IFDEF MODENT1}
  ULibExercice,
  CPProcGen,
  {$ENDIF MODENT1}
  UtilPgi
  {$IFDEF COMPTA}
  {$IFDEF CERTIFNF}
  ,uLibValidation
  {$ENDIF}
  {$IFDEF NETEXPERT}
  {$IFDEF CERTIFNF}
  , UAssistComsx
  {$ENDIF}
  {$ENDIF}
  {$ENDIF}
  ,UlibEcriture;

// GP le 07/09/2008 : 23158 :
{$IFDEF ENTREPRISE}
Const DevalidEcr = FALSE ;
{$ELSE}
Const DevalidEcr = TRUE ;
{$ENDIF}

procedure CloPerSynchro(DateArr : string);
var
    Cloture : TCloPer;
    Exo, DD : string;
begin
    if ctxPCL in V_PGI.PGIContexte then exit;
    Cloture := TCloPer.Create(Application);
    with Cloture do
    begin
        SourisSablier;
        Application.ProcessMessages;
        RempliComboExo;
        Exo := QUELEXODTBUD(StrToDate(DateArr));
        FExercice.Value := Exo;
        GenPerExo(FMois, Annee, Exo);
        FDateCpta1.Value := DateArr;
        DD := FormatDateTime('mmmm yyyy', StrtoDate(DateArr));
        FDateCpta1.ItemIndex := FDateCpta1.Items.IndexOf(DD);
        CloParSynchro := TRUE;
        if CtrlDate(DateArr) and MvtExotiquesOK(DateExo.Deb, DateClotNouv) then
        begin
            Application.ProcessMessages;
            LanceCloture;
            DoCloture(Exo);
            ClotureExoPrecedent (Exo);
        end;
        SourisNormale;
    end;
    Cloture.Free;
end;

procedure CloPer;
var
    Cloture : TCloPer;
    PP : THPanel;
begin
    if not _BlocageMonoPoste(True) then
        Exit;
    Cloture := TCloPer.Create(Application);
    Cloture.CtrlExterne := FALSE;
    Cloture.BValider.Hint := Cloture.MsgBleme.Mess[6];
    Cloture.PourAuditCloture := FALSE;
    Cloture.gbCloture := True;
    Cloture.CloParSynchro := FALSE;
    PP := FindInsidePanel;
    if PP = nil then
    begin
        try
            Cloture.ShowModal;
        finally
            Cloture.Free;
            _DeblocageMonoPoste(True);
        end;
        Screen.Cursor := SyncrDefault;
    end
    else
    begin
        InitInside(Cloture, PP);
        Cloture.Show;
    end;
end;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

procedure AnnuleCloPer;
var
    Cloture : TCloPer;
    PP : THPanel;
begin
    if not _BlocageMonoPoste(True) then
        Exit;
    Cloture := TCloPer.Create(Application);
    Cloture.Caption := 'Déclôture périodique';
    Cloture.HelpContext := 7737000;
    Cloture.CtrlExterne := FALSE;
    Cloture.BValider.Hint := Cloture.MsgBleme.Mess[10];
    Cloture.PourAuditCloture := FALSE;
    Cloture.gbCloture := False;
    Cloture.CloParSynchro := False;
    PP := FindInsidePanel;
    if PP = nil then
    begin
        try
            Cloture.ShowModal;
        finally
            Cloture.Free;
            _DeblocageMonoPoste(True);
        end;
        Screen.Cursor := SyncrDefault;
    end
    else
    begin
        InitInside(Cloture, PP);
        Cloture.Show;
    end;
end;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function CtrlTout(Exo : TExoDate) : Byte;
var
    Cloture : TCloPer;
begin
    Cloture := TCloPer.Create(Application);
    try
        Cloture.CtrlExterne := TRUE;
        Cloture.ResultatCtrlExterne := 0;
        Cloture.ExoExterne := Exo;
        Cloture.Caption := Cloture.MsgBleme.Mess[5];
        Cloture.PourAuditCloture := FALSE;
        Cloture.gbCloture := True;
        Cloture.CloParSynchro := FALSE;
        Cloture.ShowModal;
    finally
        Result := Cloture.ResultatCtrlExterne;
        Cloture.Free;
    end;
    Screen.Cursor := SyncrDefault;
end;

(* *)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function CtrlTo(Exo : TExoDate) : TTraitement;
var
    Cloture : TCloPer;
begin
    Cloture := TCloPer.Create(Application);
    try
        Cloture.CtrlExterne := TRUE;
        Cloture.ResultatCtrlExterne := 0;
        Cloture.ExoExterne := Exo;
        Cloture.Caption := Cloture.MsgBleme.Mess[5];
        Cloture.PourAuditCloture := TRUE;
        Cloture.gbCloture := True;
        Cloture.CloParSynchro := FALSE;
        Cloture.ShowModal;
        if Cloture.CFait then
            result := cOk
        else
            result := cPasFait;
    finally
        Cloture.Free;
    end;
    Screen.Cursor := SyncrDefault;
end;
(**)

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function DonneAnnee(LaDate : TDateTime) : Integer;
var
    a, m, j : Word;
begin
    DecodeDate(LaDate, a, m, j);
    Result := a;
end;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function DonneMois(LaDate : TDateTime) : Integer;
var
    a, m, j : Word;
begin
    DecodeDate(LaDate, a, m, j);
    Result := m;
end;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

procedure TCloPer.FormShow(Sender : TObject);
var
    StDate : string;
begin
    PopUpMenu := ADDMenuPop(PopUpMenu, '', '');
    RempliComboExo;
    if CtrlExterne then
    begin
        FExercice.Value := VH^.EnCours.Code;
        Htexte1.Visible := TRUE;
        HTexte2.Visible := TRUE;
        FDateCpta1.Visible := FALSE;
        HLabel3.Visible := FALSE;
        HLabel4.Visible := FALSE;
        Fexercice.Visible := FALSE;
    end
    else
    begin
        FExercice.Value := VH^.Entree.Code;
        DateClotActu := SQLDateClo;
        StDate := FormatdateTime('mmm yyyy', DateClotActu);
        HDernClo.Caption := MsgBleme.Mess[9] + ' ' + StDate;
        HDernClo.Visible := TRUE;
    end;
    RefreshMemo;
  (*
  NbLigne:=0 ;
  Resume.Lines.Clear ;
  Panel1.Height:=Resume.Top+Resume.Height+1 ;
  HPB.Top:=Panel1.Top+Panel1.Height ; ClientHeight:=HPB.Top+HPB.Height ;
  *)
    CFait := False;
end;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

procedure TCloPer.FExerciceChange(Sender : TObject);
begin
    FExercice.Enabled := False;
    GenPerExo(FMois, Annee);
end;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

procedure TCloPer.RempliComboExo;
var
    Q : TQuery;
    Sql : string;
begin
    if CtrlExterne then
        Sql := 'SELECT EX_EXERCICE,EX_LIBELLE FROM EXERCICE WHERE EX_EXERCICE="' +
            VH^.EnCours.Code + '"'
    else
        Sql := 'SELECT EX_EXERCICE,EX_LIBELLE FROM EXERCICE WHERE EX_EXERCICE="' +
            VH^.Entree.Code + '"';
    Q := OpenSql(Sql, True);
    FExercice.Values.Clear;
    FExercice.Items.Clear;
    while not Q.EOF do
    begin
        FExercice.Values.Add(Q.FindField('EX_EXERCICE').AsString);
        FExercice.Items.Add(Q.FindField('EX_LIBELLE').AsString);
        Q.Next;
    end;
    Ferme(Q);
end;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

procedure TCloPer.GenPerExo(var DebMois, LAnnee : Word;Exo : string = '');
{ Création des périodes comptables pour un exercice donné }
var
    i : integer;
    Annee, pMois, NbMois : Word;
    annee_cloture,mois_cloture,jour_cloture:word;
    annee_encours,mois_encours,jour_encours:word;
    Q : TQuery;
    DD : TdateTime;
    St : string;

begin
    NbMois := 0;
    FDateCpta1.Clear;
    FDateCpta2.clear;
    FDateCptaExo.clear;
    // YMO 24/01/2006 Réinitialisation des valeurs
    FDateCpta1.Values.Clear ;
    FDateCpta2.Values.Clear ;
    FDateCptaExo.Values.Clear ;
    if Exo <> '' then
        Q := OpenSQL('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE' +
            ' WHERE EX_EXERCICE="' + Exo + '"', TRUE)
    else
        Q := OpenSQL('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE' +
            ' WHERE EX_EXERCICE="' + FExercice.Value + '"', TRUE);
    DateExo.Deb := Q.FindField('EX_DATEDEBUT').asDateTime;
    DateExo.Fin := Q.FindField('EX_DATEFIN').asDateTime;
    NombrePerExo(DateExo, pMois, Annee, NbMois);
    DebMois := pMois;
    LAnnee := Annee;
    DD := DateExo.Deb;

    decodedate(VH^.DateCloturePer,annee_cloture,mois_cloture,jour_cloture); //SG6 15/11/04 FQ 14888


    for i := 0 to NbMois - 1 do
    begin
        //On affiche que les périodes pouvant être cloturé ou décloturé (suivant le cas) suivant la date de dernière cloture (SG6 15/11/04 FQ 14888)
        decodedate(PlusMois(DD, i),annee_encours,mois_encours,jour_encours); //SG6 15/11/04 FQ 14888
        St := FormatDateTime('mmmm yyyy', PlusMois(DD, i));
        FDateCptaExo.Items.Add(St);
        FDateCptaExo.Values.Add(DateToStr(PlusMois(DD, i)));
        if (gbCloture and (annee_cloture<annee_encours)) or (gbCloture and (mois_encours>mois_cloture) and (annee_cloture=annee_encours)) or (not(gbCloture) and (mois_encours<=mois_cloture) and (annee_encours=annee_cloture)) or ((not gbCloture) and (annee_encours < annee_cloture)) then //SG6 15/11/04 FQ 14888
        begin
          St := FormatDateTime('mmmm yyyy', PlusMois(DD, i));
      //  FDateCpta1.Items[i]:=St ; FDateCpta2.Items[i]:=St;
          FDateCpta1.Items.Add(St);
          FDateCpta1.Values.Add(DateToStr(PlusMois(DD, i)));
          FDateCpta2.Items.Add(St);
        end;
    end;

    FDateCpta1.ItemIndex := 0;
    FDateCpta2.ItemIndex := NbMois - 1;
//    FDateCptaExo.ItemIndex := 0;
    FDateCpta1Change(FDateCpta1) ;
    Ferme(Q);
end;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 29/09/2004
Modifié le ... : 29/09/2004
Description .. : - LG - 29/09/2004 - controle que la date de cloture n'est
Suite ........ : pas supp a la date de fin de l'exo ( pour les exo decale )
Mots clefs ... : 
*****************************************************************}
{function TCloPer.DonneJDate(Periodes : THValComBoBox;DebOuFin : Boolean;Mo, An : word) : TDateTime;
var
    a, m, j : Word;

begin
    if DebOuFin then
    begin
        if Periodes.ItemIndex + Mo <= 12 then
            DonnejDate := EncodeDate(An, Periodes.ItemIndex + Mo, 01)
        else
            DonnejDate := EncodeDate(An + 1, Periodes.ItemIndex + Mo - 12, 01);
    end
    else
    begin
        if Periodes.ItemIndex + Mo <= 12 then
        begin
            DecodeDate(FINDEMOIS(EncodeDate(An, Periodes.ItemIndex + Mo, 01)), a, m,
                j);
            DonnejDate := EncodeDate(An, Periodes.ItemIndex + Mo, j);
        end
        else
        begin
            //SG6 16.03.05 FQ 15501 Correction de l'encodate de la date an + 1 et non an
            DecodeDate(FINDEMOIS(EncodeDate(An + 1, Periodes.ItemIndex + Mo - 12, 01)), a,
                m, j);
            DonnejDate := EncodeDate(An + 1, Periodes.ItemIndex + Mo - 12, j);
        end;
    end;


end;  }

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function TCloPer.SQLDateClo : TDateTime;
begin
    Result := GetParamSocSecur('SO_DATECLOTUREPER', iDate1900);
end;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 18/09/2007
Modifié le ... :   /  /    
Description .. : - LG - 18/09/2007 - FB 21120 - correction pour les exercice 
Suite ........ : decalés
Mots clefs ... : 
*****************************************************************}
function TCloPer.CtrlDate(DateArr : string = '') : Boolean;
begin

    if gbCloture then
       // DateClotNouv := DonneJDate(FDateCpta1, False, FMois, Annee)
          DateClotNouv := FinDeMois( StrToDate(FDateCpta1.Value) ) // FQ 17232 SBO 16/01/2006
    else
    // DateClotNouv : Dernier jour du dernier mois clôturé
        if DateArr <> '' then // ajout me pour les synchro SX
           DateClotNouv := StrToDate(DateArr)
        else
           DateClotNouv := StrToDate(FDateCpta1.Value) - 1;

    if DateClotNouv > DateExo.Fin then
     DateClotNouv := DateExo.Fin ;


    DateClotActu := SQLDateCLo;
    Result := (DateClotActu < DateClotNouv);

    if CloParSynchro then exit; // ajout me pour ne pas demander la question
    if (not Result) and gbCloture then
        Confirmation.Execute(1, '', ''); // La période que vous avez choisie est antérieure à la date de clôture
end;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function TCloPer.MvtExotiquesOK(DAte1, Date2 : TDateTime) : Boolean;
var
    Q : TQuery;
    NatEcr : string[5];
    i, Reponse : Integer;
    Question, Abonne : Boolean;
    LastGene, NextGene : TDateTime;
    Sep, Arr : String3;

begin
    Result := False;
    Question := False;
    Abonne := False;
    NatEcr := 'SUPR';
    if PourAuditCLoture then
        NatEcr := 'SPR';
    for i := 1 to Length(NatEcr) do
    begin
        Q := OpenSQL('SELECT G_GENERAL FROM GENERAUX WHERE ' +
            'EXISTS (SELECT E_GENERAL FROM ECRITURE WHERE' +
            '               E_EXERCICE="' + FExercice.Value + '" ' +
            '           AND E_DATECOMPTABLE<="' + USDateTime(Date2) + '" ' +
            '           AND E_DATECOMPTABLE>="' + USDateTime(Date1) + '" ' +
            '           AND E_QUALIFPIECE="' + NatEcr[i] + '")', True);
        if not Q.EOF then
        begin
            Question := True;
            Inc(NbLigne);
            Resume.Lines.Add(MsgBleme.Mess[i - 1]);
        end;
        Ferme(Q);
    end;
    i := 5;
    Q := OpenSQL('SELECT CB_DATEDERNGENERE,CB_SEPAREPAR,CB_ARRONDI FROM CONTABON',
        TRUE);
    while not (Q.EOF or Abonne) do
    begin
        LastGene := Q.FindField('CB_DATEDERNGENERE').AsDateTime;
        if LastGene <= Date2 then
        begin
            Sep := Q.FindField('CB_SEPAREPAR').AsString;
            Arr := Q.FindField('CB_ARRONDI').AsString;
            NextGene := ProchaineDate(LastGene, Sep, Arr);
      //ShowMessage(USDateTime(NextGene)+' < '+USDateTime(DateClotNouv)) ;
            if NextGene <= Date2 then
                Abonne := True;
        end;
        Q.Next;
    end;
    if Abonne then
    begin
        Question := True;
        Inc(NbLigne);
        Resume.Lines.Add(MsgBleme.Mess[i - 1]);
    end;
    Ferme(Q);

    if CloParSynchro then
    begin
        Result := TRUE;
        exit;
    end; // ajout me pour ne pas demander la question

    if Question then
    begin
    //GP  Resume.Height:=18*NbLigne ;
    //GP  Panel1.Height:=Resume.Top+Resume.Height+1 ;
    //GP  HPB.Top:=Panel1.Top+Panel1.Height ;
    //GP  ClientHeight:=HPB.Top+HPB.Height ;
        Reponse := mrYes;
        if PourAuditCloture then
            Confirmation.Execute(5, '', '')
        else
            Reponse := Confirmation.Execute(0, '', '');
        case Reponse of
            mrYes :
                begin
                    Application.ProcessMessages;
                    Result := True;
                end;
            mrNo :
                begin
                    Screen.Cursor := SynCrDefault;
                end;
        end;
    end
    else
    begin
        if PourAuditCloture then
            Confirmation.Execute(4, '', '');
        Result := True;
    end;
  // Rony Fiche de bug n° 1991
    if not PourAuditCloture then
        RefreshMemo;
end;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

procedure TCloPer.LanceCloture;
begin
  SetParamSoc('SO_DATECLOTUREPER', DateClotNouv);
  VH^.DateCloturePer := DateClotNouv;
end;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

procedure TCloPer.LanceDecloture;
begin
  {JP 15/06/06 : FQ 17767 : Il arrive que le premier mois du premier exercice ne soit pas le 01 : d'où
                 le DebutDeMois pour être sûr que le paramsoc soit mis à jour avec le dernier jour du
                 mois précédent.}
  SetParamSoc('SO_DATECLOTUREPER', DebutDeMois(StrToDate(FDateCpta1.Value)) - 1);
  VH^.DateCloturePer := DebutDeMois(StrToDate(FDateCpta1.Value)) - 1;
end;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

procedure TCloPer.BFermeClick(Sender : TObject);
begin
    if (IsInside(self)) then CloseInsidePanel(self);
    Close;
end;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

procedure TCloPer.BValiderClick(Sender : TObject);
var
    Reponse : Integer;
    StDate : string;
    {$IFDEF NETEXPERT}
    IsNetEXpert : Boolean;
    {$ENDIF}
    {$IFDEF CERTIFNF}
    BlocNote : hTStringList;
    JalSess : integer;
    {$ENDIF}
begin               
{$IFNDEF CERTIFNF}
    if (not GbCloture)and (GetParamSocSecur('SO_CPCONFORMEBOI', False))
    then //06/12/2006 YMO Norme NF 203 sur la dévalidation
    begin
      PGIInfo('Pour la conformité stricte avec la norme NF 203 (et le BOI du 24/01/2006) cette fonction n''est plus disponible',Caption);
      Exit;
    end;   
{$ENDIF}
    {$IFNDEF TT}
    {$IFDEF NETEXPERT}
    if not InterrogeUserService (IsNetEXpert) then
    begin
        if IsNetExpert then
        begin
            PGIInfo('Il y a des utilisateurs connectés sur le dossier Business Line, cette fonction n''est plus disponible',Caption);
            Exit;
        end;
    end;
    {$ENDIF}
    {$ENDIF}
    if FDateCpta1.Value='' then  //SG6 15/11/04 FQ 14888
    begin
      { b md 22/09/2005 FQ 15211 }
      (*if gbCloture then HShowMessage('45;Erreur;Vous devez sélectionné une période à clôturer !;O;O;O;','','')
      else HShowMessage('45;Erreur;Vous devez sélectionné une période à déclôturer !;O;O;O;','','');*)
      if gbCloture then HShowMessage('45;Erreur;Vous devez sélectionner une période à clôturer !;E;O;O;O;','','')
      else HShowMessage('45;Erreur;Vous devez sélectionner une période à déclôturer !;E;O;O;O;','','');
      { e md }
      exit;
    end;                         //SG6 15/11/04 FQ 14888

    //YMO 11/08/2006 Norme NF/BOI Message non bloquant ecritures non validées
    If CtrlDate
    and (ExisteSQL('SELECT E_VALIDE FROM ECRITURE WHERE E_VALIDE="-"'
          +  ' AND E_EXERCICE="'       + FExercice.Value
          + '" AND E_DATECOMPTABLE>="' + USDateTime(DateExo.Deb)
          + '" AND E_DATECOMPTABLE<="' + USDateTime(DateClotNouv)
          + '" AND E_QUALIFPIECE="N" AND E_ECRANOUVEAU="N" '))
    and (Confirmation.Execute(11, '', '') <> mrYes) then
          Exit;

    if PourAuditCloture then
        RefreshMemo;

    if CtrlExterne then
    begin
        if not PourAuditCloture then
            Reponse := Confirmation.Execute(6, '', '')
        else
            Reponse := mrYes;

        EnableControls(Self, False);
        SourisSablier;

        case Reponse of
            mrYes :
                begin
                    HPatience.Caption := MsgBleme.Mess[7];
                    HPatience.Visible := TRUE;
                    Application.ProcessMessages;
                    CFait := True;
                    if MvtExotiquesOK(ExoExterne.Deb, ExoExterne.Fin) then
                    else
                        ResultatCtrlExterne := 1;
                    HPatience.Visible := FALSE;
                end;
            mrNo : Screen.Cursor := SynCrDefault;
        end;
        EnableControls(Self, TRUE);
        SourisNormale;
    end
    else
    begin
    // On effectue une clôture
        if gbCloture then
        begin
            Reponse := Confirmation.Execute(6, '', ''); // Confirmez-vous le traitement de clôture périodique?

            case Reponse of
                mrYes :
                    begin
                        EnableControls(Self, False);
                        SourisSablier;

                        HPatience.Caption := MsgBleme.Mess[7]; // Contrôle en cours. Veuillez patienter...
                        HPatience.Visible := TRUE;
                        Application.ProcessMessages;
                        if CtrlDate and MvtExotiquesOK(DateExo.Deb, DateClotNouv) then
                        begin
                            HPatience.Caption := MsgBleme.Mess[8]; // Traitement en cours. Veuillez patienter...
                            Application.ProcessMessages;

                            LanceCloture;
                            ClotureExoPrecedent;
                            DoCloture;

                            {$IFDEF NETEXPERT}
                            {$IFDEF COMPTA}
                            {$IFDEF CERTIFNF}
                            // si archivage coché
                              if BARCHIVE.Checked then
                                    EnvoiExportParDate (DateToStr(DateExo.Deb), DateToStr(DateClotActu), DateExo.Code,  'CLO');
                            {$ENDIF}
                            {$ENDIF}
                            {$ENDIF}

                            DateClotActu := SQLDateClo;
                            StDate := FormatdateTime('mmm yyyy', DateClotActu);
                            {$IFNDEF TT}
                            {$IFDEF NETEXPERT}
                              // Synchronisation avec BL
                              if gbCloture and (Reponse = mrYes) and IsNetExpert then
                                  SynchroOutBL ('P', DateTostr(DateClotActu), stDate1900);
                            {$ENDIF}
                            {$ENDIF}
                            // GCO - 18/09/2006
                            { BVE 29.08.07 : Mise en place d'un nouveau tracage }
{$IFNDEF CERTIFNF}
                            CPEnregistreLog('CLOPERIODE ' + StDate); 
{$ELSE}
{$IFDEF COMPTA}

                            BlocNote := HTStringList.Create ;
                            BlocNote.Add('CLOPERIODE ' + StDate);
                            BlocNote.Add('CHECKSUM=' + ChecksumValidation);
                            JalSess := CPEnregistreJalEvent('CPE','Clôture périodique',BlocNote);
                            // On met a jour l'enregistrement dans CPJALVALIDATION
                            SetJalEvent(SessValidation,JalSess);

{$ENDIF}
{$ENDIF}

                            HDernClo.Caption := MsgBleme.Mess[9] + ' ' + StDate; // Dernière clôture périodique :
                            Confirmation.Execute(3, '', ''); // Le traitement s'est correctement terminé.

                        end;
                        HPatience.Visible := FALSE;
                        Application.ProcessMessages;

                        EnableControls(Self, TRUE);
                        SourisNormale;
                    end;
                mrNo : Screen.Cursor := SynCrDefault;
            end;
        end
    // On effectue une déclôture
        else
        begin
            Reponse := Confirmation.Execute(7, '', ''); // Confirmez-vous le traitement de déclôture périodique?

            case Reponse of
                mrYes :
                    begin
                        EnableControls(Self, False);
                        SourisSablier;

                        HPatience.Caption := MsgBleme.Mess[7]; // Contrôle en cours. Veuillez patienter...
                        HPatience.Visible := TRUE;
                        Application.ProcessMessages;
                        if not CtrlDate then
                        begin
                            HPatience.Caption := MsgBleme.Mess[8]; // Traitement en cours. Veuillez patienter...
                            Application.ProcessMessages;

                            LanceDecloture;
                            DeclotureExoSuivant;
                            DoDecloture;
                            DateClotActu := SQLDateClo;
                            StDate := FormatdateTime('mmm yyyy', DateClotActu);
                            HDernClo.Caption := MsgBleme.Mess[9] + ' ' + StDate; // Dernière clôture périodique :
                            Confirmation.Execute(8, '', ''); // Le traitement s'est correctement terminé.

                            // GCO - 18/09/2006
                            { BVE 29.08.07 : Mise en place d'un nouveau tracage   }
{$IFNDEF CERTIFNF}
                            CPEnregistreLog('ANNULCLOPERIODE ' + FormatdateTime('mmm yyyy', PlusMois(DateClotActu,1)));
{$ELSE}
                            CPEnregistreJalEvent('CPE','Clôture périodique','ANNULCLOPERIODE ' + FormatdateTime('mmm yyyy', PlusMois(DateClotActu,1)));
{$ENDIF}
                        end;
                        HPatience.Visible := FALSE;
                        Application.ProcessMessages;

                        EnableControls(Self, TRUE);
                        SourisNormale;
                    end;
                mrNo : Screen.Cursor := SynCrDefault;
            end;
        end;
    end;

    GenPerExo(FMois, Annee); //SG6 15/11/04 FQ 14888

  // GCO et GPIO 21/10/2003 FB 12942
  //If Not CtrlExterne Then
  //  Close ;
end;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

procedure TCloPer.RefreshMemo;
begin
    NbLigne := 0;
    Resume.Lines.Clear;
end;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

procedure TCloPer.FormClose(Sender : TObject;var Action : TCloseAction);
begin
    if IsInside(Self) then
    begin
        _DeblocageMonoPoste(True);
{$IFDEF EAGLCLIENT}
{$ELSE}
        if Parent is THPanel then
            Action := caFree;
{$ENDIF}
    end;
end;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

procedure TCloPer.FormCreate(Sender : TObject);
begin
  (*
  HorzScrollBar.Range:=0 ;HorzScrollBar.Visible:=FALSE ;
  VertScrollBar.Range:=0 ;VertScrollBar.Visible:=FALSE ;
  ClientHeight:=HPB.Top+HPB.Height ; ClientWidth:=HPB.Left+HPB.Width ;
  *)
{$IFNDEF CERTIFNF}
BARCHIVE.visible := FALSE;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

procedure TCloPer.BAideClick(Sender : TObject);
begin
    CallHelpTopic(Self);
end;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Clôture l'exercice précédent

procedure TCloPer.ClotureExoPrecedent (Exo :string=''); // ajout me exo
var
    Q : TQuery;
    DateExercice : TExoDate;
    szPeriode : string[24];
    PremMois, PremAnnee, NombreMois : Word;
    i : Integer;
    szJalValide : string;
{$IFDEF CERTIFNF}
    Temp : String;
{$ENDIF}
begin
  // On se trouve sur le premier exercice ouvert : Sort
    if (Exo ='') and (VH^.Encours.Code = VH^.Entree.Code) then Exit;
    if (Exo <> '') and (VH^.Encours.Code = Exo) then exit;

  // Exercice
    Q := OpenSQL('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE EX_ETATCPTA="OUV" AND EX_EXERCICE="' + VH^.Encours.Code + '"', TRUE);
    if Q.EOf then
    begin
        Ferme(Q);
        Exit;
    end;
    DateExercice.Deb := Q.FindField('EX_DATEDEBUT').asDateTime;
    DateExercice.Fin := Q.FindField('EX_DATEFIN').asDateTime;
    Ferme(Q);
    NombrePerExo(DateExercice, PremMois, PremAnnee, NombreMois);

    szPeriode := '------------------------';
    for i := 1 to NombreMois do
        szPeriode[i] := 'X';
    ExecuteSQL('UPDATE EXERCICE SET EX_VALIDEE="' + szPeriode + '" WHERE EX_EXERCICE="' + VH^.Encours.Code + '"');

  // Ecritures                
     {  BVE 23.08.07
     Mise en place d'une nouvelle validation (NF)}
{$IFDEF COMPTA}
{$IFDEF CERTIFNF}
     Temp := ValidationEcriture('E_EXERCICE="' + VH^.Encours.Code + '" ' +
        'AND E_DATECOMPTABLE>="' + USDateTime(DateExercice.Deb) + '" ' +
        'AND E_DATECOMPTABLE<="' + USDateTime(DateExercice.Fin) + '" ' +
        'AND E_QUALIFPIECE="N" AND E_ECRANOUVEAU="N" ');
     // On traite la variable de retour :
     try
        SessValidation := StrToInt(ReadTokenST(Temp));
        ChecksumValidation := Temp; 
     except
        SessValidation := 0;
        ChecksumValidation := '';
     end;
{$ELSE}
     ExecuteSQL('UPDATE ECRITURE SET E_VALIDE="X" WHERE E_EXERCICE="' + VH^.Encours.Code + '" ' +
        'AND E_DATECOMPTABLE>="' + USDateTime(DateExercice.Deb) + '" ' +
        'AND E_DATECOMPTABLE<="' + USDateTime(DateExercice.Fin) + '" ' +
        'AND E_QUALIFPIECE="N" AND E_ECRANOUVEAU="N" ');
{$ENDIF}  
{$ELSE}
     ExecuteSQL('UPDATE ECRITURE SET E_VALIDE="X" WHERE E_EXERCICE="' + VH^.Encours.Code + '" ' +
        'AND E_DATECOMPTABLE>="' + USDateTime(DateExercice.Deb) + '" ' +
        'AND E_DATECOMPTABLE<="' + USDateTime(DateExercice.Fin) + '" ' +
        'AND E_QUALIFPIECE="N" AND E_ECRANOUVEAU="N" ');
{$ENDIF}

  // mise a jour des ventilations // FQ16762 SBO 04/10/2005 Prise en compte des ventilations à la validation
    ExecuteSQL('UPDATE ANALYTIQ SET Y_VALIDE="X" WHERE Y_EXERCICE="' + VH^.Encours.Code + '" ' +
        'AND Y_DATECOMPTABLE>="' + USDateTime(DateExercice.Deb) + '" ' +
        'AND Y_DATECOMPTABLE<="' + USDateTime(DateExercice.Fin) + '" ' +
        'AND Y_QUALIFPIECE="N" AND Y_TYPEANALYTIQUE="-" AND Y_ECRANOUVEAU="N" ');

  // Journaux
    if VH^.Entree.Code = VH^.Encours.Code then
        szJalValide := 'J_VALIDEEN1'
    else
        szJalValide := 'J_VALIDEEN';
    ExecuteSQL('UPDATE JOURNAL SET ' + szJalValide + '="' + szPeriode + '" WHERE J_NATUREJAL<>"ANO" ' +
        'AND J_NATUREJAL<>"CLO" AND J_NATUREJAL<>"ANA" AND J_NATUREJAL<>"ODA" ');
end;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

procedure TCloPer.DeclotureExoSuivant;
var
    Q : TQuery;
    DateExercice : TExoDate;
    szPeriode : string[24];
    szJalValide : string;
begin
  // On se trouve sur le dernier exercice ouvert : Sort
    if (VH^.Suivant.Code = VH^.Entree.Code) then Exit;

  // Exercice
    Q := OpenSQL('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE EX_ETATCPTA="OUV" AND EX_EXERCICE="' + VH^.Suivant.Code + '"', TRUE);
    if Q.EOf then
    begin
        Ferme(Q);
        Exit;
    end;
    DateExercice.Deb := Q.FindField('EX_DATEDEBUT').asDateTime;
    DateExercice.Fin := Q.FindField('EX_DATEFIN').asDateTime;
    Ferme(Q);
    szPeriode := '------------------------';
    ExecuteSQL('UPDATE EXERCICE SET EX_VALIDEE="' + szPeriode + '" WHERE EX_EXERCICE="' + VH^.Suivant.Code + '"');

  // Ecritures
  {  BVE 23.08.07
     On ne valide plus les ecritures sur la décloture (NF)       }
{$IFNDEF CERTIFNF}
    If DevalidEcr Then
    ExecuteSQL('UPDATE ECRITURE SET E_VALIDE="-" WHERE E_EXERCICE="' + VH^.Suivant.Code + '" ' +
        'AND E_DATECOMPTABLE>="' + USDateTime(DateExercice.Deb) + '" ' +
        'AND E_DATECOMPTABLE<="' + USDateTime(DateExercice.Fin) + '" ' +
        'AND E_QUALIFPIECE="N" AND E_ECRANOUVEAU="N" ');
{$ENDIF}

  // mise a jour des ventilations // FQ16762 SBO 04/10/2005 Prise en compte des ventilations à la validation
    If DevalidEcr Then
    ExecuteSQL('UPDATE ANALYTIQ SET Y_VALIDE="-" WHERE Y_EXERCICE="' + VH^.Encours.Code + '" ' +
        'AND Y_DATECOMPTABLE>="' + USDateTime(DateExercice.Deb) + '" ' +
        'AND Y_DATECOMPTABLE<="' + USDateTime(DateExercice.Fin) + '" ' +
        'AND Y_QUALIFPIECE="N" AND Y_TYPEANALYTIQUE="-" AND Y_ECRANOUVEAU="N" ');

  // Journaux
    if VH^.Entree.Code = VH^.Suivant.Code then
        szJalValide := 'J_VALIDEEN'
    else
        szJalValide := 'J_VALIDEEN1';
    ExecuteSQL('UPDATE JOURNAL SET ' + szJalValide + '="' + szPeriode + '" WHERE J_NATUREJAL<>"ANO" ' +
        'AND J_NATUREJAL<>"CLO" AND J_NATUREJAL<>"ANA" AND J_NATUREJAL<>"ODA" ');
end;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

procedure TCloPer.DoCloture(Exo : string = '');
var
    Q : TQuery;
    DateDeb : TDateTime;
    i : Integer;
    szPeriode : string[24];
    szJalValide : string;
    CodeExo : string;     
{$IFDEF CERTIFNF}
    Temp : String;
{$ENDIF}
begin
    if Exo <> '' then
        CodeExo := Exo
    else
        CodeExo := FExercice.Value;

    if DateClotNouv > DateClotActu then
    begin
        if VH^.Entree.Code = VH^.Encours.Code then
            szJalValide := 'J_VALIDEEN'
        else
            szJalValide := 'J_VALIDEEN1';

    // Exercice
        Q := OpenSQL('SELECT EX_DATEDEBUT FROM EXERCICE WHERE EX_EXERCICE="' + CodeExo + '"', TRUE);
        DateDeb := Q.FindField('EX_DATEDEBUT').asDateTime;

        szPeriode := '------------------------';
        for i := 1 to 24 do
        begin
// GP 23077
//            if (i <= FDateCpta1.ItemIndex + 1) then
            if (i <= FDateCptaExo.ItemIndex + 1) then
                szPeriode[i] := 'X'
            else
                szPeriode[i] := '-';
        end;
        ExecuteSQL('UPDATE EXERCICE SET EX_VALIDEE="' + szPeriode + '" WHERE EX_EXERCICE="' + CodeExo + '"');

    // Ecritures comptables  
  {  BVE 23.08.07
     Mise en place d'une nouvelle validation (NF) }
{$IFDEF COMPTA}
{$IFDEF CERTIFNF}
        Temp := ValidationEcriture('E_EXERCICE="' + CodeExo + '" ' +
            'AND E_DATECOMPTABLE>="' + USDateTime(DateDeb) + '" ' +
            'AND E_DATECOMPTABLE<="' + USDateTime(DateClotNouv) + '" ' +
            'AND E_QUALIFPIECE="N" AND E_ECRANOUVEAU="N" ');      
        // On traite la variable de retour :
        try
           SessValidation := StrToInt(ReadTokenST(Temp));
           ChecksumValidation := Temp;
        except
           SessValidation := 0;
           ChecksumValidation := '';
        end;
{$ELSE}
        ExecuteSQL('UPDATE ECRITURE SET E_VALIDE="X" WHERE E_EXERCICE="' + CodeExo + '" ' +
            'AND E_DATECOMPTABLE>="' + USDateTime(DateDeb) + '" ' +
            'AND E_DATECOMPTABLE<="' + USDateTime(DateClotNouv) + '" ' +
            'AND E_QUALIFPIECE="N" AND E_ECRANOUVEAU="N" ');
{$ENDIF}   
{$ELSE}
        ExecuteSQL('UPDATE ECRITURE SET E_VALIDE="X" WHERE E_EXERCICE="' + CodeExo + '" ' +
            'AND E_DATECOMPTABLE>="' + USDateTime(DateDeb) + '" ' +
            'AND E_DATECOMPTABLE<="' + USDateTime(DateClotNouv) + '" ' +
            'AND E_QUALIFPIECE="N" AND E_ECRANOUVEAU="N" ');
{$ENDIF}


    // Journaux
        ExecuteSQL('UPDATE JOURNAL SET ' + szJalValide + '="' + szPeriode + '" WHERE J_NATUREJAL<>"ANO" ' +
            'AND J_NATUREJAL<>"CLO" AND J_NATUREJAL<>"ANA" AND J_NATUREJAL<>"ODA" ');

    // Recharge les tablettes
        AvertirMultiTable('TTJOURNAL');
    end;
end;

procedure TCloPer.DoDecloture;
var
{$IFNDEF CERTIFNF}
    Q : TQuery;
    DateFin : TDateTime;
{$ENDIF}
    i : Integer;
    szPeriode : string[24];
    szJalValide : string;
begin
    if DateClotNouv < DateClotActu then
    begin
        if VH^.Entree.Code = VH^.Encours.Code then
            szJalValide := 'J_VALIDEEN'
        else
            szJalValide := 'J_VALIDEEN1';

{$IFNDEF CERTIFNF}
        // Exercice
        Q := OpenSQL('SELECT EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE="' + FExercice.Value + '"', TRUE);
        DateFin := Q.FindField('EX_DATEFIN').asDateTime;
        Ferme(Q);
{$ENDIF}

        szPeriode := '------------------------';
        for i := 1 to 24 do
        begin
// GP 23077
//            if (i <= FDateCpta1.ItemIndex) then
            if (i <= FDateCptaExo.ItemIndex) then
                szPeriode[i] := 'X'
            else
                szPeriode[i] := '-';
        end;
        ExecuteSQL('UPDATE EXERCICE SET EX_VALIDEE="' + szPeriode + '" WHERE EX_EXERCICE="' + VH^.Suivant.Code + '"');

    // Ecritures comptables
    { BVE 23.08.07
    On ne valide plus les écritures sur la décloture  }
{$IFNDEF CERTIFNF}
    If DevalidEcr Then
    ExecuteSQL('UPDATE ECRITURE SET E_VALIDE="-" WHERE E_EXERCICE="' + FExercice.Value + '" ' +
          'AND E_DATECOMPTABLE>="' + USDateTime(DateClotNouv + 1) + '" ' +
          'AND E_DATECOMPTABLE<="' + USDateTime(DateFin) + '" ' +
          'AND E_QUALIFPIECE="N" AND E_ECRANOUVEAU="N" ');
{$ENDIF}

    // Journaux
        ExecuteSQL('UPDATE JOURNAL SET ' + szJalValide + '="' + szPeriode + '" WHERE J_NATUREJAL<>"ANO" ' +
            'AND J_NATUREJAL<>"CLO" AND J_NATUREJAL<>"ANA" AND J_NATUREJAL<>"ODA" ');

    // Recharge les tablettes
        AvertirMultiTable('TTJOURNAL');
    end;
end;

procedure TCloPer.BArchiveClick(Sender: TObject);
begin
{$IFDEF CERTIFNF}
 if not BARCHIVE.Checked then
 begin
       If PGIAsk ('En référence au BOI 13 L-1-06 N° 12 du 24 Janvier 2006 Paragraphe 95,'+#10#13+
       ' il est préconisé d''archiver l''ensemble des informations dont la conservation est obligatoire.'+#10#13+
       ' Vous avez choisi de ne pas effectuer un archivage automatique de ces données, la responsabilité '+#10#13+
       ' de le réaliser manuellement vous incombe. Confirmez-vous ce choix ?')  <> mrYes then
               BARCHIVE.Checked := TRUE;
 end;
{$ENDIF}
end;

procedure TCloPer.FDateCpta1Change(Sender: TObject);
Var St : String ;
begin
St:=FDateCpta1.Value ;
If St<>'' Then FDateCptaExo.value:=St ;
end;

end.
