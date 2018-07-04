unit AssistCiblageVersOperation;

interface

uses
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fe_Main,
{$else}
     eMul,
     MainEagl,
{$ENDIF}
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  utomOperation,            //affecteCodeAuto
  assist, ComCtrls, HSysMenu, hmsgbox, StdCtrls, ExtCtrls, HPanel, HTB97, UtilGC,
  Hctrls, Mask, DBCtrls, HDB, ColMemo, UTob, HRichEdt, HRichOLE, ParamSoc, KPMGUtil,
  Grids, utilTOM, lookup
  , EntRT, UtilRT
  ;

  procedure Assist_CiblageVersOperation ( CodeCiblage : String);


type
  TFAssistCiblageVersOperation = class(TFAssist)
    Operation: TTabSheet;
    Actiongenerique: TTabSheet;
    TROP_OPERATION: THLabel;
    ROP_OPERATION: THCritMaskEdit;
    ROP_LIBELLE: THCritMaskEdit;
    TROP_LIBELLE: THLabel;
    ROP_DATEDEBUT: THCritMaskEdit;
    TROP_DATEDEBUT: THLabel;
    ROP_DATEFIN: THCritMaskEdit;
    TROP_DATEFIN: THLabel;
    ROP_BUDGET: THCritMaskEdit;
    TROP_BUDGET: THLabel;
    ROP_COUT: THCritMaskEdit;
    TROP_COUT: THLabel;
    TROP_BLOCNOTE: THLabel;
    ROP_BLOCNOTE: THRichEditOLE;
    RAG_TYPEACTION: THValComboBox;
    TRAG_TYPEACTION: THLabel;
    RAG_LIBELLE: THCritMaskEdit;
    TRAG_LIBELLE: THLabel;
    RAG_DATEACTION: THCritMaskEdit;
    TRAG_DATEACTION: THLabel;
    RAG_DATEECHEANCE: THCritMaskEdit;
    TRAG_DATEECHEANCE: THLabel;
    RAG_COUTACTION: THCritMaskEdit;
    TRAG_COUTACTION: THLabel;
    Recap: TTabSheet;
    HRECAP: THListBox;
    TCodeCiblage: THCritMaskEdit;
    GAction: THGrid;
    BAjoute: TToolbarButton97;
    BSupprime: TToolbarButton97;
    BValider: TToolbarButton97;
    BAnnule: TToolbarButton97;
    bMonter: TToolbarButton97;
    bDescendre: TToolbarButton97;
    TMODELE_OPERATION: THLabel;
    RAG_TABLELIBRE1: THValComboBox;
    RAG_TABLELIBRE2: THValComboBox;
    RAG_TABLELIBRE3: THValComboBox;
    TRAG_TABLELIBRE1: THLabel;
    TRAG_TABLELIBRE2: THLabel;      //TJA FQ 10694 nom de champ incorrect
    TRAG_TABLELIBRE3: THLabel;
    MODELE_OPERATION: THCritMaskEdit;
    ROP_OBJETOPE: THValComboBox;
    TROP_OBJETOPE: THLabel;
    RAG_ETATACTION: THValComboBox;
    TRAG_ETATACTION: THLabel;
    TsInfoCompl: TTabSheet;
    ROP_ROPTABLELIBRE1: THValComboBox;
    TROP_ROPTABLELIBRE1: THLabel;
    ROP_ROPTABLELIBRE2: THValComboBox;
    ROP_ROPTABLELIBRE3: THValComboBox;
    ROP_ROPTABLELIBRE4: THValComboBox;
    ROP_ROPTABLELIBRE5: THValComboBox;
    TROP_ROPTABLELIBRE2: THLabel;
    TROP_ROPTABLELIBRE3: THLabel;
    TROP_ROPTABLELIBRE4: THLabel;
    TROP_ROPTABLELIBRE5: THLabel;
    Contacts: TTabSheet;
    TC_FONCTIONCODEE: THLabel;
    C_FONCTIONCODEE: THMultiValComboBox;
    C_PRINCIPAL: THCheckbox;
    C_PUBLIPOSTAGE: THCheckbox;
    C_FERME: THCheckbox;
    TC_LIBRECONTACT1: THLabel;
    C_LIBRECONTACT1: THMultiValComboBox;
    HLabel1: THLabel;
    TC_LIBRECONTACT2: THLabel;
    C_LIBRECONTACT2: THMultiValComboBox;
    TC_LIBRECONTACT3: THLabel;
    TC_LIBRECONTACT4: THLabel;
    C_LIBRECONTACT3: THMultiValComboBox;
    C_LIBRECONTACT4: THMultiValComboBox;
    TC_LIBRECONTACT5: THLabel;
    C_LIBRECONTACT5: THMultiValComboBox;
    TC_LIBRECONTACT6: THLabel;
    C_LIBRECONTACT6: THMultiValComboBox;
    TC_LIBRECONTACT7: THLabel;
    C_LIBRECONTACT7: THMultiValComboBox;
    TC_LIBRECONTACT8: THLabel;
    C_LIBRECONTACT8: THMultiValComboBox;
    TC_LIBRECONTACT9: THLabel;
    C_LIBRECONTACT9: THMultiValComboBox;
    TC_LIBRECONTACTA: THLabel;
    C_LIBRECONTACTA: THMultiValComboBox;
    CC_LIBRECONTACT1: THValComboBox;
    CC_LIBRECONTACT2: THValComboBox;
    CC_LIBRECONTACT3: THValComboBox;
    CC_LIBRECONTACT4: THValComboBox;
    CC_LIBRECONTACT5: THValComboBox;
    CC_LIBRECONTACT6: THValComboBox;
    CC_LIBRECONTACT7: THValComboBox;
    CC_LIBRECONTACT8: THValComboBox;
    CC_LIBRECONTACT9: THValComboBox;
    CC_LIBRECONTACTA: THValComboBox;
    CC_FONCTIONCODEE: THValComboBox;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure RechActionGenerique;
    procedure BAjouteClick(Sender: TObject);
    procedure BAnnuleClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure GActionClick(Sender: TObject);
    procedure BSupprimeClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure bMonterClick(Sender: TObject);
    procedure bDescendreClick(Sender: TObject);
    procedure MODELE_OPERATIONElipsisClick(Sender: TObject);
    procedure CorrespCiblageOperation ();
    procedure RAG_TYPEACTIONExit(Sender: TObject);



  private
    { Déclarations privées }
    TobCiblage          : TOB;
    TobOperation        : TOB;
    TobActionGenerique  : TOB;
    LesColonnes         : String;
    EnSaisie            : Boolean;
    TobLienCorr         : Tob;
    bNumAuto            : Boolean;
    bTablesLibres       : Boolean;    //FQ 10731
    StrJoinContact      : String;     //FQ 10541
    WhereContact        : String;     //FQ 10541

    procedure InitForme;
    function FormateCombo(StrCombo:string):string; //FQ 10541

  public
        { Déclarations publiques }
  end;

var
  FAssistCiblageVersOperation: TFAssistCiblageVersOperation;

implementation

{$R *.DFM}


procedure Assist_CiblageVersOperation ( CodeCiblage : String);
var
  F                     : TFAssistCiblageVersOperation;

begin
  F                     := TFAssistCiblageVersOperation.Create(Application);
  F.TCodeCiblage.Text   := CodeCiblage;
  //Recheche si le code opération est automatique ou bien attribué par le code ciblage
  F.bNumAuto            := GetParamSocSecur('SO_RTNUMOPERAUTO', True, True);
  if F.bNumAuto then
    F.ROP_OPERATION.Enabled := False
  else
    F.ROP_OPERATION.Text    := CodeCiblage;

  Try
    F.ShowModal;
  Finally
    F.Free;
  end;

end;


procedure TFAssistCiblageVersOperation.FormCreate(Sender: TObject);

begin
  inherited;

  TobCiblage            := Tob.Create('CIBLAGE', Nil, -1);
  TobOperation          := Tob.Create('OPERATIONS', Nil, -1);
  TobActionGenerique    := Tob.Create('ACTIONSGENERIQUES', Nil, -1);

end;


procedure TFAssistCiblageVersOperation.FormShow(Sender: TObject);
begin
  inherited;
   InitForme;
end;




procedure TFAssistCiblageVersOperation.InitForme;
var
  i                     : integer;
  LaCombo               : THValComboBox;
  QQ                    : TQUERY;
  StrReq, StrLib        : String;
begin
   FAssistCiblageVersOperation.Caption := 'Assistant ciblage vers opération';

  // FAssistCiblageVersOperation.bFin.Enabled := False; //bouton 'Démarrer' non actif

   ROP_DATEDEBUT.Text   := DateToStr(date) ;
   RAG_DATEACTION.Text  := ROP_DATEDEBUT.Text;

   TobCiblage.SelectDB('"'+ TCodeCiblage.Text +'"', nil, False);
   ROP_LIBELLE.Text     := TobCiblage.getvalue('RCB_LIBELLE');
   ROP_BLOCNOTE.Text    := TobCiblage.getValue('RCB_MEMO');

   //Affectation du combo pour actions génériques
   RAG_TYPEACTION.Items.Clear;
   RAG_TYPEACTION.Values.Clear;

  RAG_TYPEACTION.DataType   := 'RTTYPEACTION';

  LesColonnes               := 'RAG_TYPEACTION;RAG_LIBELLE';

  for i := 1 to 3 do
    ChangeLibre2('TRAG_TABLELIBRE' + IntToStr(i), TForm(Self));

  RAG_ETATACTION.DataType   := 'RTETATACTION';

  //on cache le champ avec elipsis pour les modèle d'opération
  MODELE_OPERATION.Visible  := False;
  TMODELE_OPERATION.Visible := False;

  // Les tables libres opération
  GCMAJChampLibre(TForm(Self), False, 'COMBO', 'ROP_ROPTABLELIBRE', 5, '_');

  //recherche de correspondances ressources <-> ciblage <-> opérations
  TobLienCorr           := Tob.Create('correspondances ressources', nil, -1);
  CorrespCiblageOperation;
  for i := 0 to TobLienCorr.Detail.Count -1 do
  begin
    LaCombo             := THValComboBox(TOMTOFGetControl(Tform(Self), TobLienCorr.Detail[i].GetValue('CHOPERATION')));
    Lacombo.Value       := TobCiblage.GetValue(TobLienCorr.Detail[i].GetValue('CHCIBLAGE'));
  end;

  // FQ 10731 TJA 28/02/2008
  // recherche s'il y a au moins une table libre utilisée, sinon on gère le "non arrêt" dans l'onglet tables libres
  bTablesLibres         := False;
  for i := 1 to 5 do
  begin
    LaCombo             := THValComboBox(TForm(Self).FindComponent('ROP_ROPTABLELIBRE' + IntToStr(i)));
    if LaCombo.Visible = True then
    begin
      bTablesLibres     := True;
      Break;
    end;
  end;

  //TJA 02/07/2008
  //FQ 10867  boutons du 3eme onglet au look V2008
  BAjoute.GlobalIndexImage    := 'Z0053_S16G1';
  BSupprime.GlobalIndexImage  := 'Z0005_S16G1';
  BValider.GlobalIndexImage   := 'Z0003_S16G2';
  BAnnule.GlobalIndexImage    := 'Z0021_S16G1';
  bMonter.GlobalIndexImage    := 'Z2202_S16G1';
  bDescendre.GlobalIndexImage := 'Z1373_S16G1';

  // STR 03/07/2008 FQ 10541
  if TobCiblage.GetString('RCB_TYPECIBLAGE')='002' then
  begin  // ciblage par contact
    C_LIBRECONTACT1.visible := true;
    C_LIBRECONTACT2.visible := true;
    C_LIBRECONTACT3.visible := true;
    C_LIBRECONTACT4.visible := true;
    C_LIBRECONTACT5.visible := true;
    C_LIBRECONTACT6.visible := true;
    C_LIBRECONTACT7.visible := true;
    C_LIBRECONTACT8.visible := true;
    C_LIBRECONTACT9.visible := true;
    C_LIBRECONTACTA.visible := true;
    C_FONCTIONCODEE.visible := true;
    CC_LIBRECONTACT1.visible := false;
    CC_LIBRECONTACT2.visible := false;
    CC_LIBRECONTACT3.visible := false;
    CC_LIBRECONTACT4.visible := false;
    CC_LIBRECONTACT5.visible := false;
    CC_LIBRECONTACT6.visible := false;
    CC_LIBRECONTACT7.visible := false;
    CC_LIBRECONTACT8.visible := false;
    CC_LIBRECONTACT9.visible := false;
    CC_LIBRECONTACTA.visible := false;
    CC_FONCTIONCODEE.visible := false;
  end else begin // ciblage par tiers
    C_LIBRECONTACT1.visible := false;
    C_LIBRECONTACT2.visible := false;
    C_LIBRECONTACT3.visible := false;
    C_LIBRECONTACT4.visible := false;
    C_LIBRECONTACT5.visible := false;
    C_LIBRECONTACT6.visible := false;
    C_LIBRECONTACT7.visible := false;
    C_LIBRECONTACT8.visible := false;
    C_LIBRECONTACT9.visible := false;
    C_LIBRECONTACTA.visible := false;
    C_FONCTIONCODEE.visible := false;
    CC_LIBRECONTACT1.visible := true;
    CC_LIBRECONTACT2.visible := true;
    CC_LIBRECONTACT3.visible := true;
    CC_LIBRECONTACT4.visible := true;
    CC_LIBRECONTACT5.visible := true;
    CC_LIBRECONTACT6.visible := true;
    CC_LIBRECONTACT7.visible := true;
    CC_LIBRECONTACT8.visible := true;
    CC_LIBRECONTACT9.visible := true;
    CC_LIBRECONTACTA.visible := true;
    CC_FONCTIONCODEE.visible := true;
  end;
  // vérif si tables libres contacts paramétrées + récup libellé associé
  for i := 1 to 10 do
  begin
    StrReq := 'select cc_libelle from choixcod where cc_type="ZLC" and cc_code="BT';
    if i<10 then
      StrREq := StrReq+IntToStr(i)+'"'
    else
      StrREq := StrReq+'A"';
    StrLib := '.-';
    try
      QQ := OpenSql(StrReq,true);
      if Not QQ.EOf then
        StrLib := QQ.FindField('cc_libelle').AsString;
    Finally
      Ferme(QQ);
    end;
    if Copy(StrLib,1,2)='.-' then
    begin
      Case i of
        1 : begin TC_LIBRECONTACT1.visible := false; C_LIBRECONTACT1.visible := false; CC_LIBRECONTACT1.visible := false; end;
        2 : begin TC_LIBRECONTACT2.visible := false; C_LIBRECONTACT2.visible := false; CC_LIBRECONTACT2.visible := false; end;
        3 : begin TC_LIBRECONTACT3.visible := false; C_LIBRECONTACT3.visible := false; CC_LIBRECONTACT3.visible := false; end;
        4 : begin TC_LIBRECONTACT4.visible := false; C_LIBRECONTACT4.visible := false; CC_LIBRECONTACT4.visible := false; end;
        5 : begin TC_LIBRECONTACT5.visible := false; C_LIBRECONTACT5.visible := false; CC_LIBRECONTACT5.visible := false; end;
        6 : begin TC_LIBRECONTACT6.visible := false; C_LIBRECONTACT6.visible := false; CC_LIBRECONTACT6.visible := false; end;
        7 : begin TC_LIBRECONTACT7.visible := false; C_LIBRECONTACT7.visible := false; CC_LIBRECONTACT7.visible := false; end;
        8 : begin TC_LIBRECONTACT8.visible := false; C_LIBRECONTACT8.visible := false; CC_LIBRECONTACT8.visible := false; end;
        9 : begin TC_LIBRECONTACT9.visible := false; C_LIBRECONTACT9.visible := false; CC_LIBRECONTACT9.visible := false; end;
        10: begin TC_LIBRECONTACTA.visible := false; C_LIBRECONTACTA.visible := false; CC_LIBRECONTACTA.visible := false; end;
      end;
    end else
    begin
      Case i of
        1 : TC_LIBRECONTACT1.caption := StrLib;
        2 : TC_LIBRECONTACT2.caption := StrLib;
        3 : TC_LIBRECONTACT3.caption := StrLib;
        4 : TC_LIBRECONTACT4.caption := StrLib;
        5 : TC_LIBRECONTACT5.caption := StrLib;
        6 : TC_LIBRECONTACT6.caption := StrLib;
        7 : TC_LIBRECONTACT7.caption := StrLib;
        8 : TC_LIBRECONTACT8.caption := StrLib;
        9 : TC_LIBRECONTACT9.caption := StrLib;
        10: TC_LIBRECONTACTA.caption := StrLib;
      end;
    end;
  end;
  //
end;


//FQ 10541
function TFAssistCiblageVersOperation.FormateCombo(StrCombo:string):string;
var Wcombo, StrTmp : string;
begin
  result := '';
  Wcombo := StrCombo;
  StrTmp := ReadTokenSt(Wcombo);
  While StrTmp<>'' do
  begin
    if result='' then Result:='"' else result:=result+',"';
    result := result+StrTmp+'"';
    StrTmp := ReadTokenSt(Wcombo);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 20/06/2006
Modifié le ... :   /  /
Description .. : Gestion du bouton 'Suivant'
Mots clefs ... :
*****************************************************************}
procedure TFAssistCiblageVersOperation.bSuivantClick(Sender: TObject);
Var
   Onglet : TTabSheet;
   StrOnglet : String;
   StrRecap : String;
   NbAction : integer;
   ii : integer;
   StrJoinSuite, WhereIn : String;
begin

//procédures avant de changer d'onglet

Onglet := P.ActivePage;
StrOnglet := Onglet.Name;

//gestion des éléments page 'opération'
if StrOnglet = 'Operation' then
begin
   if (not GetParamSocSecur('SO_RTNUMOPERAUTO', False, true)) and (ROP_OPERATION.Text = '') then
   begin
      PGIError('Il faut un code opération', Caption);
      ROP_OPERATION.SetFocus;
      exit;
   end;

  //TJA FQ 10669 on test l'existance de l'opération dans le cas de la saisie du code
  if (not bNumAuto) and (ExisteSQL('SELECT 1 FROM OPERATIONS WHERE ROP_OPERATION = "' + ROP_OPERATION.Text + '"')) then
  begin
    PGIError('Ce code opération existe déjà !#10#13 Il faut un code opération valide.', Caption);
    ROP_OPERATION.Text  := '';
    ROP_OPERATION.SetFocus;
    exit;
  end;

   if ROP_LIBELLE.Text = '' then
   begin
      PGIError('Il faut un libellé');
      ROP_LIBELLE.SetFocus;
      exit;
   end;

   RAG_DATEACTION.Text := ROP_DATEDEBUT.Text;   //la date action est au moins la même que celle de l'opération

end;

if StrOnglet = 'Actiongenerique' then
begin
   //si aucune action générique
   if TobActionGenerique.Detail.Count = 0 then
   begin
      PGIError('Il faut au moins une action générique');
      exit;
   end else begin // Controle type d'action fermé
     for ii := 0 to TobActionGenerique.Detail.Count-1 do
     begin
       if ExisteSql('select rpa_ferme from paractions where rpa_ferme="X" and rpa_typeaction="'+TobActionGenerique.Detail[ii].GetString('RAG_TYPEACTION')+'"') then
       begin
        PGIError('Le type d''action est fermé pour l'' action générique : '+TobActionGenerique.Detail[ii].GetString('RAG_TYPEACTION'));
        exit;
       end;
     end;
   end;
   MODELE_OPERATION.Visible := False;
   TMODELE_OPERATION.Visible := False;
end;


inherited;     //on change d'onglet

Onglet := P.ActivePage;
StrOnglet := Onglet.Name;

// Onglet Tables libres
if StrOnglet = 'TsInfoCompl' then
  if not bTablesLibres then
    bSuivantClick(Self);




//gestion de la page 'Actiongenerique'
if StrOnglet = 'Actiongenerique' then
begin
   MODELE_OPERATION.Visible := true;
   TMODELE_OPERATION.Visible := true;
   RechActionGenerique;
   if GAction.RowCount > 2 then
   begin
      BSupprime.Enabled := True;
      Bajoute.Enabled := True;
      BValider.Enabled := True;
      BAnnule.Enabled := True;
   end;
end;


//Gestion de la page 'Recap'
If StrOnglet = 'Recap' then
begin

    Bfin.Enabled := True;
    HRECAP.Items.Clear;
    StrRecap := 'OPERATION';
    HRECAP.Items.Add(StrRecap);
    StrRecap := '------------------------------';
    HRECAP.Items.Add(StrRecap);
    If GetParamSocSecur('SO_RTNUMOPERAUTO', False, true) then
      StrRecap := 'Code : Attribution automatique'
    else
      StrRecap := 'Code : '+ROP_OPERATION.Text;
    HRECAP.Items.Add(StrRecap);
    StrRecap := 'Libellé : '+ROP_LIBELLE.Text;
    HRECAP.Items.Add(StrRecap);
    StrRecap := 'Date de début : '+ROP_DATEDEBUT.Text;
    HRECAP.Items.Add(StrRecap);
    StrRecap := 'Date de Fin : '+ROP_DATEFIN.Text;
    HRECAP.Items.Add(StrRecap);
    if ROP_BUDGET.Text <> '' then
    begin
      StrRecap := 'Budget : '+ROP_BUDGET.text+' €';
      HRECAP.Items.Add(StrRecap);
    end;
    if ROP_COUT.Text <> '' then
    begin
      StrRecap := 'Coût : '+ROP_COUT.text+' €';
      HRECAP.Items.Add(StrRecap);
    end;
    StrRecap := '==============================';
    HRECAP.Items.Add(StrRecap);

    StrRecap := 'ACTION GENERIQUE';
    HRECAP.Items.Add(StrRecap);
    StrRecap := '------------------------------';
    HRECAP.Items.Add(StrRecap);
    for NbAction := 0 to TobActionGenerique.Detail.Count-1 do
    begin
      StrRecap := 'Action : '+TobActionGenerique.Detail[NbAction].Getvalue('RAG_TYPEACTION');
      HRECAP.Items.Add(StrRecap);
      StrRecap := 'Libellé : '+TobActionGenerique.Detail[NbAction].Getvalue('RAG_LIBELLE');
      HRECAP.Items.Add(StrRecap);
      StrRecap := 'Date début : '+DateToStr(TobActionGenerique.Detail[NbAction].Getvalue('RAG_DATEACTION'));
      HRECAP.Items.Add(StrRecap);
      StrRecap := 'Date échéance : '+DateTOStr(TobActionGenerique.Detail[NbAction].Getvalue('RAG_DATEECHEANCE'));
      HRECAP.Items.Add(StrRecap);
      if RAG_COUTACTION.Text <> '' then
      begin
         StrRecap := 'Coût : '+IntToStr(TobActionGenerique.Detail[NbAction].Getvalue('RAG_COUTACTION'))+' €';
         HRECAP.Items.Add(StrRecap);
      end;
      StrRecap := '';
      HRECAP.Items.Add(StrRecap);
    end;
    // FQ 10541 - STR
    StrJoinContact := '';
    StrJoinSuite := '';
    WhereContact := '';
    StrRecap := '==============================';
    HRECAP.Items.Add(StrRecap);
    StrRecap := 'CONTACTS';
    HRECAP.Items.Add(StrRecap);
    StrRecap := '------------------------------';
    HRECAP.Items.Add(StrRecap);
    if C_FONCTIONCODEE.Text<>'' then
    begin
      StrRecap := 'Fonction : '+C_FONCTIONCODEE.Text;
      HRECAP.Items.Add(StrRecap);
    end;
    if C_PRINCIPAL.state<>cbGrayed then
    begin
      if C_PRINCIPAL.Checked then
        StrRecap := 'Contact principal : OUI'
      else
        StrRecap := 'Contact principal : NON';
      HRECAP.Items.Add(StrRecap);
    end;
    if C_PUBLIPOSTAGE.state<>cbGrayed then
    begin
      if C_PUBLIPOSTAGE.Checked then
        StrRecap := 'Publipostage : OUI'
      else
        StrRecap := 'Publipostage : NON';
      HRECAP.Items.Add(StrRecap);
    end;
    if C_FERME.state<>cbGrayed then
    begin
      if C_FERME.Checked then
        StrRecap := 'Contact fermé : OUI'
      else
        StrRecap := 'Contact fermé : NON';
      HRECAP.Items.Add(StrRecap);
    end;
    if (C_LIBRECONTACT1.Text<>'') then
    begin
      StrRecap := 'Table libre contact 1 : '+C_LIBRECONTACT1.Text;
      HRECAP.Items.Add(StrRecap);
      WhereIn := FormateCombo(C_LIBRECONTACT1.Text);
      WhereContact := WhereContact+' AND C_LIBRECONTACT1 IN ('+WhereIn+') ';
    end;
    if (CC_LIBRECONTACT1.Text<>'')and(CC_LIBRECONTACT1.value<>CC_LIBRECONTACT1.VideString) then
    begin
      StrRecap := 'Table libre contact 1 : '+CC_LIBRECONTACT1.Text;
      HRECAP.Items.Add(StrRecap);
      StrJoinSuite := StrJoinSuite+' AND C_LIBRECONTACT1="'+CC_LIBRECONTACT1.value+'" ';
    end;
    if (C_LIBRECONTACT2.Text<>'') then
    begin
      StrRecap := 'Table libre contact 2 : '+C_LIBRECONTACT2.Text;
      HRECAP.Items.Add(StrRecap);
      WhereIn := FormateCombo(C_LIBRECONTACT2.Text);
      WhereContact := WhereContact+' AND C_LIBRECONTACT2 IN ('+WhereIn+') ';
    end;
    if (CC_LIBRECONTACT2.Text<>'')and(CC_LIBRECONTACT2.value<>CC_LIBRECONTACT2.VideString) then
    begin
      StrRecap := 'Table libre contact 2 : '+CC_LIBRECONTACT2.Text;
      HRECAP.Items.Add(StrRecap);
      StrJoinSuite := StrJoinSuite+' AND C_LIBRECONTACT2="'+CC_LIBRECONTACT2.value+'" ';
    end;
    if (C_LIBRECONTACT3.Text<>'') then
    begin
      StrRecap := 'Table libre contact 3 : '+C_LIBRECONTACT3.Text;
      HRECAP.Items.Add(StrRecap);
      WhereIn := FormateCombo(C_LIBRECONTACT3.Text);
      WhereContact := WhereContact+' AND C_LIBRECONTACT3 IN ('+WhereIn+') ';
    end;
    if (CC_LIBRECONTACT3.Text<>'')and(CC_LIBRECONTACT3.value<>CC_LIBRECONTACT3.VideString) then
    begin
      StrRecap := 'Table libre contact 3 : '+CC_LIBRECONTACT3.Text;
      HRECAP.Items.Add(StrRecap);
      StrJoinSuite := StrJoinSuite+' AND C_LIBRECONTACT3="'+CC_LIBRECONTACT3.value+'" ';
    end;
    if (C_LIBRECONTACT4.Text<>'') then
    begin
      StrRecap := 'Table libre contact 4 : '+C_LIBRECONTACT4.Text;
      HRECAP.Items.Add(StrRecap);
      WhereIn := FormateCombo(C_LIBRECONTACT4.Text);
      WhereContact := WhereContact+' AND C_LIBRECONTACT4 IN ('+WhereIn+') ';
    end;
    if (CC_LIBRECONTACT4.Text<>'')and(CC_LIBRECONTACT4.value<>CC_LIBRECONTACT4.VideString) then
    begin
      StrRecap := 'Table libre contact 4 : '+CC_LIBRECONTACT4.Text;
      HRECAP.Items.Add(StrRecap);
      StrJoinSuite := StrJoinSuite+' AND C_LIBRECONTACT4="'+CC_LIBRECONTACT4.value+'" ';
    end;
    if (C_LIBRECONTACT5.Text<>'') then
    begin
      StrRecap := 'Table libre contact 5 : '+C_LIBRECONTACT5.Text;
      HRECAP.Items.Add(StrRecap);
      WhereIn := FormateCombo(C_LIBRECONTACT5.Text);
      WhereContact := WhereContact+' AND C_LIBRECONTACT5 IN ('+WhereIn+') ';
    end;
    if (CC_LIBRECONTACT5.Text<>'')and(CC_LIBRECONTACT5.value<>CC_LIBRECONTACT5.VideString) then
    begin
      StrRecap := 'Table libre contact 5 : '+CC_LIBRECONTACT5.Text;
      HRECAP.Items.Add(StrRecap);
      StrJoinSuite := StrJoinSuite+' AND C_LIBRECONTACT5="'+CC_LIBRECONTACT5.value+'" ';
    end;
    if (C_LIBRECONTACT6.Text<>'') then
    begin
      StrRecap := 'Table libre contact 6 : '+C_LIBRECONTACT6.Text;
      HRECAP.Items.Add(StrRecap);
      WhereIn := FormateCombo(C_LIBRECONTACT6.Text);
      WhereContact := WhereContact+' AND C_LIBRECONTACT6 IN ('+WhereIn+') ';
    end;
    if (CC_LIBRECONTACT6.Text<>'')and(CC_LIBRECONTACT6.value<>CC_LIBRECONTACT6.VideString) then
    begin
      StrRecap := 'Table libre contact 6 : '+CC_LIBRECONTACT6.Text;
      HRECAP.Items.Add(StrRecap);
      StrJoinSuite := StrJoinSuite+' AND C_LIBRECONTACT6="'+CC_LIBRECONTACT6.value+'" ';
    end;
    if (C_LIBRECONTACT7.Text<>'') then
    begin
      StrRecap := 'Table libre contact 7 : '+C_LIBRECONTACT7.Text;
      HRECAP.Items.Add(StrRecap);
      WhereIn := FormateCombo(C_LIBRECONTACT7.Text);
      WhereContact := WhereContact+' AND C_LIBRECONTACT7 IN ('+WhereIn+') ';
    end;
    if (CC_LIBRECONTACT7.Text<>'')and(CC_LIBRECONTACT7.value<>CC_LIBRECONTACT7.VideString) then
    begin
      StrRecap := 'Table libre contact 7 : '+CC_LIBRECONTACT7.Text;
      HRECAP.Items.Add(StrRecap);
      StrJoinSuite := StrJoinSuite+' AND C_LIBRECONTACT7="'+CC_LIBRECONTACT7.value+'" ';
    end;
    if (C_LIBRECONTACT8.Text<>'') then
    begin
      StrRecap := 'Table libre contact 8 : '+C_LIBRECONTACT8.Text;
      HRECAP.Items.Add(StrRecap);
      WhereIn := FormateCombo(C_LIBRECONTACT8.Text);
      WhereContact := WhereContact+' AND C_LIBRECONTACT8 IN ('+WhereIn+') ';
    end;
    if (CC_LIBRECONTACT8.Text<>'')and(CC_LIBRECONTACT8.value<>CC_LIBRECONTACT8.VideString) then
    begin
      StrRecap := 'Table libre contact 8 : '+CC_LIBRECONTACT8.Text;
      HRECAP.Items.Add(StrRecap);
      StrJoinSuite := StrJoinSuite+' AND C_LIBRECONTACT8="'+CC_LIBRECONTACT8.value+'" ';
    end;
    if (C_LIBRECONTACT9.Text<>'') then
    begin
      StrRecap := 'Table libre contact 9 : '+C_LIBRECONTACT9.Text;
      HRECAP.Items.Add(StrRecap);
      WhereIn := FormateCombo(C_LIBRECONTACT9.Text);
      WhereContact := WhereContact+' AND C_LIBRECONTACT9 IN ('+WhereIn+') ';
    end;
    if (CC_LIBRECONTACT9.Text<>'')and(CC_LIBRECONTACT9.value<>CC_LIBRECONTACT9.VideString) then
    begin
      StrRecap := 'Table libre contact 9 : '+CC_LIBRECONTACT9.Text;
      HRECAP.Items.Add(StrRecap);
      StrJoinSuite := StrJoinSuite+' AND C_LIBRECONTACT9="'+CC_LIBRECONTACT9.value+'" ';
    end;
    if (C_LIBRECONTACTA.Text<>'') then
    begin
      StrRecap := 'Table libre contact 10 : '+C_LIBRECONTACTA.Text;
      HRECAP.Items.Add(StrRecap);
      WhereIn := FormateCombo(C_LIBRECONTACTA.Text);
      WhereContact := WhereContact+' AND C_LIBRECONTACTA IN ('+WhereIn+') ';
    end;
    if (CC_LIBRECONTACTA.Text<>'')and(CC_LIBRECONTACTA.value<>CC_LIBRECONTACTA.VideString) then
    begin
      StrRecap := 'Table libre contact 10 : '+CC_LIBRECONTACTA.Text;
      HRECAP.Items.Add(StrRecap);
      StrJoinSuite := StrJoinSuite+' AND C_LIBRECONTACTA="'+CC_LIBRECONTACTA.value+'" ';
    end;
    // Si on est en selection multiple sur les contacts
    //if WhereContact<>'' then
    if TobCiblage.GetSTring('RCB_TYPECIBLAGE')='002' then
    begin
      StrJoinContact := ' LEFT JOIN CONTACT ON C_TYPECONTACT="T" AND C_NATUREAUXI="CLI" AND C_AUXILIAIRE=RVB_CONTACT AND C_NUMEROCONTACT=RVB_NUMEROCONTACT ';
      // critère "fonction"
      if C_FONCTIONCODEE.Text<>'' then
      begin
        WhereIn := FormateCombo(C_FONCTIONCODEE.Text);
        WhereContact := WhereContact+' AND C_FONCTIONCODEE in ('+WhereIn+') ';
      end;
      // critère "contact principal"
      if C_PRINCIPAL.state<>cbGrayed then
        if C_PRINCIPAL.checked then
          WhereContact := WhereContact+' AND C_PRINCIPAL="X" '
        else
          WhereContact := WhereContact+' AND C_PRINCIPAL="-" ';
      // critère "Publipostage"
      if C_PUBLIPOSTAGE.state<>cbGrayed then
        if C_PUBLIPOSTAGE.checked then
          WhereContact := WhereContact+' AND C_PUBLIPOSTAGE="X" '
        else
          WhereContact := WhereContact+' AND C_PUBLIPOSTAGE="-" ';
      // critère "Fermé"
      if C_FERME.state<>cbGrayed then
        if C_FERME.checked then
          WhereContact := WhereContact+' AND C_FERME="X" '
        else
          WhereContact := WhereContact+' AND C_FERME="-" ';
      //
    end
    // sinon sélection simple sur les contacts
    else //if StrJoinSuite<>'' then
    begin
      // critère "fonction"
      if (C_FONCTIONCODEE.Text<>'') and(CC_FONCTIONCODEE.value<>CC_FONCTIONCODEE.VideString) then
      begin
        StrJoinContact := StrJoinContact+' AND C_FONCTIONCODEE="'+C_FONCTIONCODEE.Text+'" ';
      end;
      // critère "contact principal"
      if C_PRINCIPAL.state<>cbGrayed then
        if C_PRINCIPAL.checked then
          StrJoinContact := StrJoinContact+' AND C_PRINCIPAL="X" '
        else
          StrJoinContact := StrJoinContact+' AND C_PRINCIPAL="-" ';
      // critère "Publipostage"
      if C_PUBLIPOSTAGE.state<>cbGrayed then
        if C_PUBLIPOSTAGE.checked then
          StrJoinContact := StrJoinContact+' AND C_PUBLIPOSTAGE="X" '
        else
          StrJoinContact := StrJoinContact+' AND C_PUBLIPOSTAGE="-" ';
      // critère "Fermé"
      if C_FERME.state<>cbGrayed then
        if C_FERME.checked then
          StrJoinContact := StrJoinContact+' AND C_FERME="X" '
        else
          StrJoinContact := StrJoinContact+' AND C_FERME="-" ';
      if (StrJoinContact<>'')or(StrJoinSuite<>'') then
        StrJoinContact := ' LEFT JOIN CONTACT ON C_TYPECONTACT="T" AND C_NATUREAUXI="CLI" AND C_TIERS=RVB_TIERS '+StrJoinSuite+StrJoinContact;
      //
    end;

end;


end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 20/06/2006
Modifié le ... :   /  /
Description .. : Gestion du bouton 'Precedent'
Mots clefs ... :
*****************************************************************}
procedure TFAssistCiblageVersOperation.bPrecedentClick(Sender: TObject);
var
   Onglet : TTabSheet;
   StrOnglet : String;
begin
Onglet := P.ActivePage;
StrOnglet := Onglet.Name;


if StrOnglet = 'Recap' then
   Bfin.Enabled := False;

if StrOnglet = 'Actiongenerique' then
begin
   //si aucune action générique
   if TobActionGenerique.Detail.Count = 0 then
   begin
      PGIError('Il faut au moins une action générique');
      exit;
   end;
   MODELE_OPERATION.Visible := False;
   TMODELE_OPERATION.Visible := False;
end;


inherited;     // on change de page
Onglet := P.ActivePage;
StrOnglet := Onglet.Name;


// Onglet Tables libres
if StrOnglet = 'TsInfoCompl' then
  if not bTablesLibres then
  bPrecedentClick(Self);


if StrOnglet = 'Actiongenerique' then
begin
   RechActionGenerique;
   MODELE_OPERATION.Visible := True;
   TMODELE_OPERATION.Visible := True;
   if GAction.RowCount > 2 then
   begin
      BSupprime.Enabled := True;
      Bajoute.Enabled := True;
      BValider.Enabled := True;
      BAnnule.Enabled := True;
   end;
end;

end;


procedure TFAssistCiblageVersOperation.bFinClick(Sender: TObject);
var
  CodeOperation         : String;
  NbAction              : integer;
  CodeAuto              : Boolean;
  NumChrono             : String;

begin
Inherited;

  CodeOperation         := '';
  CodeAuto              := GetParamSocSecur('SO_RTNUMOPERAUTO', False, true);

  if CodeAuto then
    CodeOperation       := AffecteCodeAuto        //TJA 27/06/2007  Gestion du code auto avec gestion du préfixage du code user
  else
    CodeOperation       := ROP_OPERATION.Text;

  //affectation valeurs opération
  TobOperation.PutValue('ROP_OPERATION', CodeOperation);
  TobOperation.PutValue('ROP_LIBELLE', ROP_LIBELLE.Text);
  TobOperation.PutValue('ROP_PRODUITPGI', 'GRC');
  TobOperation.PutValue('ROP_DATEDEBUT', StrToDate(ROP_DATEDEBUT.Text));
  TobOperation.PutValue('ROP_DATEFIN', StrToDate(ROP_DATEFIN.Text));
  if ROP_BUDGET.Text = '' then
    TobOperation.PutValue('ROP_BUDGET', 0)
  else
    TobOperation.PutValue('ROP_BUDGET', StrToInt(ROP_BUDGET.Text));
  if ROP_COUT.Text = '' then
    TobOperation.PutValue('ROP_COUT', 0)
  else
    TobOperation.PutValue('ROP_COUT', StrToInt(ROP_COUT.Text));
  TobOperation.PutValue('ROP_OBJETOPE', '');
  TobOperation.PutValue('ROP_FERME', '-');
  TobOperation.PutValue('ROP_OBJETOPEF', '');
  TobOperation.PutValue('ROP_BLOCNOTE', RichToString(ROP_BLOCNOTE));
  TobOperation.PutValue('ROP_OBJETOPE', ROP_OBJETOPE.Value);
  TobOperation.PutValue('ROP_ROPTABLELIBRE1', ROP_ROPTABLELIBRE1.Value);
  TobOperation.PutValue('ROP_ROPTABLELIBRE2', ROP_ROPTABLELIBRE2.Value);
  TobOperation.PutValue('ROP_ROPTABLELIBRE3', ROP_ROPTABLELIBRE3.Value);
  TobOperation.PutValue('ROP_ROPTABLELIBRE4', ROP_ROPTABLELIBRE4.Value);
  TobOperation.PutValue('ROP_ROPTABLELIBRE5', ROP_ROPTABLELIBRE5.Value);


  //affectation valeurs dans les actions génériques
  for NbAction := 0 to TobActionGenerique.Detail.Count-1 do
  begin
   TobActionGenerique.Detail[NbAction].PutValue('RAG_OPERATION', CodeOperation);
   TobActionGenerique.Detail[NbAction].PutValue('RAG_NUMACTGEN', NbAction+1);
   TobActionGenerique.Detail[NbAction].PutValue('RAG_PRODUITPGI', 'GRC');
   TobActionGenerique.Detail[NbAction].PutValue('RAG_INTERVENANT', TobCiblage.GetValue('RCB_RESPONSABLE'));
    //   TobActionGenerique.Detail[NbAction].PutValue('RAG_ETATACTION', 'PRE');
   TobActionGenerique.Detail[NbAction].PutValue('RAG_BLOCNOTE', '');
  end;


  HRECAP.Items.Add('');

  //if Ciblage2Operation(TobOperation, TobActionGenerique, TCodeCiblage.Text) then
  // FQ 10541 : ajout des paramètres whereContact et StrJoinContact
  if Ciblage2Operation(TobOperation, TobActionGenerique, TCodeCiblage.Text, WhereContact, StrJoinContact) then
  begin
    if CodeAuto then
    begin
      NumChrono         := ExtraitChronoCode(CodeOperation);
      SetParamSoc('SO_RTCOMPTEUROPER', NumChrono) ;
    end;
    HRECAP.Items.Add('Opération enregistrée');
    close;
  end
  else
    HRECAP.Items.Add('Impossible de créer l''opération');

end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 20/06/2006
Modifié le ... :   /  /
Description .. : Recherche les actions génériques déjà enregistrées
Mots clefs ... :
*****************************************************************}
procedure TFAssistCiblageVersOperation.RechActionGenerique;
var
   NbAction : integer;
begin

GAction.ClearSelected;
GAction.RowCount := TobActionGenerique.Detail.Count+2;

//La tob est vide
if TobActionGenerique.Detail.Count = 0 then
begin
   RAG_TYPEACTION.Enabled := False;
   RAG_LIBELLE.Enabled := False;
   RAG_DATEACTION.Enabled := False;
   RAG_ETATACTION.Enabled := False;
   RAG_DATEECHEANCE.Enabled := False;
   RAG_COUTACTION.Enabled := False;
   RAG_TABLELIBRE1.Enabled := False;
   RAG_TABLELIBRE2.Enabled := False;
   RAG_TABLELIBRE3.Enabled := False;

   BSupprime.Enabled := False;
   BValider.Enabled := False;
   BAnnule.Enabled := False;
   BAjoute.Enabled := True;
end
else
//s'il y a au moins une action générique
begin
   BSupprime.Enabled := True;
   BAjoute.Enabled := True;
   //FQ 10871 - STR 20/10/08: on doit pouvoir modifier l'action générique même s'il n'y en a qu'une
   //BValider.Enabled := False;
   //BAnnule.Enabled := False;
   BValider.Enabled := True;
   BAnnule.Enabled := True;
   For NbAction := 0 to TobActionGenerique.Detail.Count-1 do
   begin
      TobActionGenerique.PutGridDetail(GAction,False,False,LesColonnes,True);
   end;
   //chargement des infos de la premiere action
   RAG_TYPEACTION.Value := TobActionGenerique.Detail[0].getvalue('RAG_TYPEACTION');
   RAG_LIBELLE.Text := TobActionGenerique.Detail[0].getvalue('RAG_LIBELLE');
   RAG_DATEACTION.text := DateToStr(TobActionGenerique.Detail[0].getvalue('RAG_DATEACTION'));
   RAG_DATEECHEANCE.Text := DateToStr(TobActionGenerique.Detail[0].getvalue('RAG_DATEECHEANCE'));
   RAG_COUTACTION.Text := TobActionGenerique.Detail[0].getvalue('RAG_COUTACTION');
   RAG_TABLELIBRE1.Value := TobActionGenerique.Detail[0].getvalue('RAG_TABLELIBRE1');
   RAG_TABLELIBRE2.Value := TobActionGenerique.Detail[0].getvalue('RAG_TABLELIBRE2');
   RAG_TABLELIBRE3.Value := TobActionGenerique.Detail[0].getvalue('RAG_TABLELIBRE3');
end;




end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 20/06/2006
Modifié le ... :   /  /
Description .. : Bouton de création d'une nouvelle action générique
Mots clefs ... :
*****************************************************************}
procedure TFAssistCiblageVersOperation.BAjouteClick(Sender: TObject);
begin
inherited;
  EnSaisie              := True;

  bSuivant.Enabled      := False;
  bPrecedent.Enabled    := False;

  GAction.Enabled       := False;

  BAjoute.Enabled       := False;
  BSupprime.Enabled     := False;
  BValider.Enabled      := True;
  BAnnule.Enabled       := True;
  bMonter.Enabled       := False;
  bDescendre.Enabled    := False;

  RAG_TYPEACTION.Enabled    := True;
  RAG_LIBELLE.Enabled       := True;
  RAG_DATEACTION.Enabled    := True;
  RAG_ETATACTION.Enabled    := True;
  RAG_DATEECHEANCE.Enabled  := True;
  RAG_COUTACTION.Enabled    := True;
  RAG_TABLELIBRE1.Enabled   := True;
  RAG_TABLELIBRE2.Enabled   := True;
  RAG_TABLELIBRE3.Enabled   := True;

  RAG_TYPEACTION.Text   := '';
  RAG_LIBELLE.Text      := '';
//  RAG_DATEACTION.Text   := ROP_DATEDEBUT.Text;
  RAG_DATEECHEANCE.DefaultDate  := od2099;    //FQ 10698   TJA
  RAG_ETATACTION.Value  := 'PRE';
  RAG_DATEECHEANCE.Text := ROP_DATEFIN.Text;
  RAG_COUTACTION.Text   := '';
  RAG_TABLELIBRE1.Value := '';
  RAG_TABLELIBRE2.Value := '';
  RAG_TABLELIBRE3.Value := '';




  RAG_TYPEACTION.SetFocus;

end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 20/06/2006
Modifié le ... :   /  /
Description .. : Annule la saisie d'action générique en cours
Mots clefs ... :
*****************************************************************}
procedure TFAssistCiblageVersOperation.BAnnuleClick(Sender: TObject);
var
   IndexAction : integer;
begin
inherited;
bSuivant.Enabled := True;
bPrecedent.Enabled := True;
GAction.Enabled := True;

EnSaisie := False;


BAjoute.Enabled := True;
if TobActionGenerique.Detail.Count = 0 then
begin
   BSupprime.Enabled := False;
   Bajoute.Enabled := True;
   BValider.Enabled := False;
   BAnnule.Enabled := False;
   bMonter.Enabled := True;
   bDescendre.Enabled := True;

   RAG_TYPEACTION.Enabled := False;
   RAG_LIBELLE.Enabled := False;
   RAG_DATEACTION.Enabled := False;
   RAG_ETATACTION.Enabled := False;
   RAG_DATEECHEANCE.Enabled := False;
   RAG_COUTACTION.Enabled := False;
   RAG_TABLELIBRE1.Enabled := False;
   RAG_TABLELIBRE2.Enabled := False;
   RAG_TABLELIBRE3.Enabled := False;

   RAG_TYPEACTION.Value := '';
   RAG_LIBELLE.Text := '';
   RAG_DATEACTION.Text := '';
   RAG_ETATACTION.Value := '';
   RAG_DATEECHEANCE.Text := '';
   RAG_COUTACTION.Text := '';
   RAG_TABLELIBRE1.Value := '';
   RAG_TABLELIBRE2.Value := '';
   RAG_TABLELIBRE3.value := '';
end
else
begin
   IndexAction := GAction.row-1;

   BSupprime.Enabled := True;
   Bajoute.Enabled := True;
   BValider.Enabled := True;
   BAnnule.Enabled := True;


   RAG_TYPEACTION.Value := TobActionGenerique.Detail[IndexAction].getvalue('RAG_TYPEACTION');
   RAG_LIBELLE.Text := TobActionGenerique.Detail[IndexAction].getvalue('RAG_LIBELLE');
   RAG_DATEACTION.text := DateToStr(TobActionGenerique.Detail[IndexAction].getvalue('RAG_DATEACTION'));
   RAG_ETATACTION.Value := TobActionGenerique.Detail[IndexAction].getvalue('RAG_ETATACTION');
   RAG_DATEECHEANCE.Text := DateToStr(TobActionGenerique.Detail[IndexAction].getvalue('RAG_DATEECHEANCE'));
   RAG_COUTACTION.Text := TobActionGenerique.Detail[IndexAction].getvalue('RAG_COUTACTION');
   RAG_TABLELIBRE1.Value := TobActionGenerique.Detail[IndexAction].getvalue('RAG_TABLELIBRE1');
   RAG_TABLELIBRE2.Value := TobActionGenerique.Detail[IndexAction].getvalue('RAG_TABLELIBRE2');
   RAG_TABLELIBRE3.Value := TobActionGenerique.Detail[IndexAction].getvalue('RAG_TABLELIBRE3');

end;



end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 20/06/2006
Modifié le ... :   /  /
Description .. : Gesion du bouton 'valider' d'une action générique
Mots clefs ... :
*****************************************************************}
procedure TFAssistCiblageVersOperation.BValiderClick(Sender: TObject);
Var
   NbAction : integer;
   IndexAction : integer;

   TobAc : TOB;

begin

if RAG_TYPEACTION.Value = '' then
begin
   PGIError('Il faut un type d''action');
   RAG_TYPEACTION.SetFocus;
   exit;
end;

if RAG_LIBELLE.Text = '' then
begin
   PGIError('Il faut un libellé');
   RAG_LIBELLE.SetFocus;
   exit;
end;



if EnSaisie then                 // Nouvelle Action
begin
   TobAc := Tob.Create('ACTIONSGENERIQUES', TobActionGenerique, -1);
   TobAc.PutValue('RAG_LIBELLE', RAG_LIBELLE.Text);
   TobAc.PutValue('RAG_TYPEACTION', RAG_TYPEACTION.Value);
   TobAc.PutValue('RAG_DATEACTION', StrToDate(RAG_DATEACTION.Text));
   TobAc.PutValue('RAG_ETATACTION', RAG_ETATACTION.Value);
   TobAc.PutValue('RAG_DATEECHEANCE', StrToDate(RAG_DATEECHEANCE.Text));
   if RAG_COUTACTION.Text = '' then
      TobAc.PutValue('RAG_COUTACTION', 0)
   else
      TobAc.PutValue('RAG_COUTACTION', StrToInt(RAG_COUTACTION.Text));
   Tobac.PutValue('RAG_TABLELIBRE1', RAG_TABLELIBRE1.Value);
   TobAc.PutValue('RAG_TABLELIBRE2', RAG_TABLELIBRE2.Value);
   TobAc.PutValue('RAG_TABLELIBRE3', RAG_TABLELIBRE3.Value);

   IndexAction := TobActionGenerique.Detail.Count;
end
else
begin
   IndexAction := GAction.row-1;
   TobActionGenerique.Detail[IndexAction].PutValue('RAG_LIBELLE', RAG_LIBELLE.Text);
   TobActionGenerique.Detail[IndexAction].PutValue('RAG_TYPEACTION', RAG_TYPEACTION.Value);
   TobActionGenerique.Detail[IndexAction].PutValue('RAG_DATEACTION', StrToDate(RAG_DATEACTION.Text));
   TobActionGenerique.Detail[IndexAction].PutValue('RAG_ETATACTION', RAG_ETATACTION.Value);
   TobActionGenerique.Detail[IndexAction].PutValue('RAG_DATEECHEANCE', StrToDate(RAG_DATEECHEANCE.Text));
   if RAG_COUTACTION.Text = '' then
      TobActionGenerique.Detail[IndexAction].PutValue('RAG_COUTACTION', 0)
   else
      TobActionGenerique.Detail[IndexAction].PutValue('RAG_COUTACTION', StrToInt(RAG_COUTACTION.Text));

   TobActionGenerique.Detail[IndexAction].PutValue('RAG_TABLELIBRE1', RAG_TABLELIBRE1.Value);
   TobActionGenerique.Detail[IndexAction].PutValue('RAG_TABLELIBRE2', RAG_TABLELIBRE2.Value);
   TobActionGenerique.Detail[IndexAction].PutValue('RAG_TABLELIBRE3', RAG_TABLELIBRE3.Value);

   Inc(IndexAction);

end;

For NbAction := 0 to TobActionGenerique.Detail.Count-1 do
begin
   TobActionGenerique.PutGridDetail(GAction,False,False,LesColonnes,True);
end;


bSuivant.Enabled := True;
bPrecedent.Enabled := True;
GAction.Enabled := True;

BAjoute.Enabled := True;
BSupprime.Enabled := True;
BValider.Enabled := True;
BAnnule.Enabled := True;
bMonter.Enabled := True;
bDescendre.Enabled := True;


inherited;

GAction.GotoRow(IndexAction);
EnSaisie := False;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 21/06/2006
Modifié le ... :   /  /
Description .. : Click sur la grille pour visualiser les infos d'une action
Suite ........ : générique
Mots clefs ... :
*****************************************************************}
procedure TFAssistCiblageVersOperation.GActionClick(Sender: TObject);
var
   IndexAction : integer;
   TobTypAct : TOB;
begin
inherited;

IndexAction := GAction.Row-1;
if TobActionGenerique.Detail.Count = 0 then          //il n'y a pas d'action
   exit;
RAG_TYPEACTION.Value := TobActionGenerique.Detail[IndexAction].getvalue('RAG_TYPEACTION');
// FQ 10888
//mcd conseil TobTypAct := Nil;
VH_RT.TobTypesAction.Load;
TobTypAct:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[RAG_TYPEACTION.Value,'---',0],TRUE) ;
if TobTypAct <> Nil then
begin
if (TobTypAct.GetValue('RPA_GESTDATECH') = 'X') then
  RAG_DATEECHEANCE.enabled := TRUE
else
  RAG_DATEECHEANCE.Enabled := FALSE ;
if (TobTypAct.GetValue('RPA_GESTCOUT') = 'X') then
  RAG_COUTACTION.Enabled := TRUE
else
  RAG_COUTACTION.Enabled := FALSE ;
end;
//
RAG_LIBELLE.Text := TobActionGenerique.Detail[IndexAction].getvalue('RAG_LIBELLE');
RAG_DATEACTION.text := DateToStr(TobActionGenerique.Detail[IndexAction].getvalue('RAG_DATEACTION'));
RAG_ETATACTION.Value := TobActionGenerique.Detail[IndexAction].GetValue('RAG_ETATACTION');
RAG_DATEECHEANCE.Text := DateToStr(TobActionGenerique.Detail[IndexAction].getvalue('RAG_DATEECHEANCE'));
RAG_COUTACTION.Text := TobActionGenerique.Detail[IndexAction].getvalue('RAG_COUTACTION');
RAG_TABLELIBRE1.Value := TobActionGenerique.Detail[IndexAction].getvalue('RAG_TABLELIBRE1');
RAG_TABLELIBRE2.Value := TobActionGenerique.Detail[IndexAction].getvalue('RAG_TABLELIBRE2');
RAG_TABLELIBRE3.Value := TobActionGenerique.Detail[IndexAction].GetValue('RAG_TABLELIBRE3');



end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 21/06/2006
Modifié le ... :   /  /
Description .. : Bouton Supprime une action générique
Mots clefs ... :
*****************************************************************}
procedure TFAssistCiblageVersOperation.BSupprimeClick(Sender: TObject);
var
   IndexAction : integer;
   NbAction : integer;

begin

inherited;

IndexAction := GAction.Row-1;
TobActionGenerique.Detail[IndexAction].Free;

GAction.DeleteRow(Gaction.Row);

For NbAction := 0 to TobActionGenerique.Detail.Count-1 do
begin
   TobActionGenerique.PutGridDetail(GAction,False,False,LesColonnes,True);
end;

if TobActionGenerique.Detail.Count = 0 then
begin
   BSupprime.Enabled := False;
   Bajoute.Enabled := True;
   BValider.Enabled := False;
   BAnnule.Enabled := False;

   RAG_TYPEACTION.Enabled := False;
   RAG_LIBELLE.Enabled := False;
   RAG_DATEACTION.Enabled := False;
   RAG_ETATACTION.Enabled := False;
   RAG_DATEECHEANCE.Enabled := False;
   RAG_COUTACTION.Enabled := False;
   RAG_TABLELIBRE1.Enabled := False;
   RAG_TABLELIBRE2.Enabled := False;
   RAG_TABLELIBRE3.Enabled := False;


   RAG_TYPEACTION.Value := '';
   RAG_LIBELLE.Text := '';
   RAG_DATEACTION.Text := '';
   RAG_ETATACTION.Value := '';
   RAG_DATEECHEANCE.Text := '';
   RAG_COUTACTION.Text := '';
   RAG_TABLELIBRE1.Value := '';
   RAG_TABLELIBRE2.Value := '';
   RAG_TABLELIBRE3.Value := '';

end
else
begin
   BSupprime.Enabled := True;
   Bajoute.Enabled := True;
   BValider.Enabled := True;
   BAnnule.Enabled := True;

   RAG_TYPEACTION.Value := TobActionGenerique.Detail[0].getvalue('RAG_TYPEACTION');
   RAG_LIBELLE.Text := TobActionGenerique.Detail[0].getvalue('RAG_LIBELLE');
   RAG_DATEACTION.text := DateToStr(TobActionGenerique.Detail[0].getvalue('RAG_DATEACTION'));
   RAG_ETATACTION.Value := TobActionGenerique.Detail[0].GetValue('RAG_ETATACTION');
   RAG_DATEECHEANCE.Text := DateToStr(TobActionGenerique.Detail[0].getvalue('RAG_DATEECHEANCE'));
   RAG_COUTACTION.Text := TobActionGenerique.Detail[0].getvalue('RAG_COUTACTION');
   RAG_TABLELIBRE1.Value := TobActionGenerique.Detail[0].getvalue('RAG_TABLELIBRE1');
   RAG_TABLELIBRE2.Value := TobActionGenerique.Detail[0].getvalue('RAG_TABLELIBRE2');
   RAG_TABLELIBRE3.Value := TobActionGenerique.Detail[0].getvalue('RAG_TABLELIBRE3');
   GAction.GotoRow(1);


end;



end;

procedure TFAssistCiblageVersOperation.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  //inherited;

  TobOperation.free;
  TobActionGenerique.free;
  TobLienCorr.Free;

end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 22/06/2006
Modifié le ... :   /  /
Description .. : Bouton faire remonter l'action génrique
Mots clefs ... :
*****************************************************************}
procedure TFAssistCiblageVersOperation.bMonterClick(Sender: TObject);
var
   Tobtampon : TOB;
   Nbaction : integer;
   IndexAction : integer;

begin
If TobActionGenerique.Detail.Count < 2 then     // aucune action ou juste une action
   exit;

if GAction.row = 1 then    // C'est la première action générique
   exit;

IndexAction := GAction.row;

//on mémorise l'action à remonter
Tobtampon := Tob.Create('ACTIONSGENERIQUES', nil, -1);
Tobtampon.putValue('RAG_TYPEACTION',TobActionGenerique.Detail[IndexAction-1].getvalue('RAG_TYPEACTION'));
Tobtampon.putValue('RAG_LIBELLE',TobActionGenerique.Detail[IndexAction-1].getvalue('RAG_LIBELLE'));
Tobtampon.putValue('RAG_DATEACTION',TobActionGenerique.Detail[IndexAction-1].getvalue('RAG_DATEACTION'));
Tobtampon.PutValue('RAG_ETATACTION', TobActionGenerique.Detail[IndexAction-1].GetValue('RAG_ETATACTION'));
Tobtampon.putValue('RAG_DATEECHEANCE',TobActionGenerique.Detail[IndexAction-1].getvalue('RAG_DATEECHEANCE'));
Tobtampon.putValue('RAG_COUTACTION',TobActionGenerique.Detail[IndexAction-1].getvalue('RAG_COUTACTION'));
Tobtampon.PutValue('RAG_TABLELIBRE1', TobActionGenerique.Detail[IndexAction-1].GetValue('RAG_TABLELIBRE1'));
Tobtampon.PutValue('RAG_TABLELIBRE2', TobActionGenerique.Detail[IndexAction-1].GetValue('RAG_TABLELIBRE2'));
Tobtampon.PutValue('RAG_TABLELIBRE3', TobActionGenerique.Detail[IndexAction-1].GetValue('RAG_TABLELIBRE3'));

//on descend celle au dessus
TobActionGenerique.Detail[IndexAction-1].putValue('RAG_TYPEACTION', TobActionGenerique.Detail[IndexAction-2].getvalue('RAG_TYPEACTION'));
TobActionGenerique.Detail[IndexAction-1].putValue('RAG_LIBELLE', TobActionGenerique.Detail[IndexAction-2].getvalue('RAG_LIBELLE'));
TobActionGenerique.Detail[IndexAction-1].putValue('RAG_DATEACTION', TobActionGenerique.Detail[IndexAction-2].getvalue('RAG_DATEACTION'));
TobActionGenerique.Detail[IndexAction-1].putValue('RAG_ETATACTION', TobActionGenerique.Detail[IndexAction-2].getvalue('RAG_ETATACTION'));
TobActionGenerique.Detail[IndexAction-1].putValue('RAG_DATEECHEANCE', TobActionGenerique.Detail[IndexAction-2].getvalue('RAG_DATEECHEANCE'));
TobActionGenerique.Detail[IndexAction-1].putValue('RAG_COUTACTION', TobActionGenerique.Detail[IndexAction-2].getvalue('RAG_COUTACTION'));
TobActionGenerique.Detail[IndexAction-1].putValue('RAG_TABLELIBRE1', TobActionGenerique.Detail[IndexAction-2].getvalue('RAG_TABLELIBRE1'));
TobActionGenerique.Detail[IndexAction-1].putValue('RAG_TABLELIBRE2', TobActionGenerique.Detail[IndexAction-2].getvalue('RAG_TABLELIBRE2'));
TobActionGenerique.Detail[IndexAction-1].putValue('RAG_TABLELIBRE3', TobActionGenerique.Detail[IndexAction-2].getvalue('RAG_TABLELIBRE3'));


//on affecte le tampon à l'action du dessus
TobActionGenerique.Detail[IndexAction-2].putValue('RAG_TYPEACTION', Tobtampon.getvalue('RAG_TYPEACTION'));
TobActionGenerique.Detail[IndexAction-2].putValue('RAG_LIBELLE', Tobtampon.getvalue('RAG_LIBELLE'));
TobActionGenerique.Detail[IndexAction-2].putValue('RAG_DATEACTION', Tobtampon.getvalue('RAG_DATEACTION'));
TobActionGenerique.Detail[IndexAction-2].putValue('RAG_ETATACTION', Tobtampon.getvalue('RAG_ETATACTION'));
TobActionGenerique.Detail[IndexAction-2].putValue('RAG_DATEECHEANCE', Tobtampon.getvalue('RAG_DATEECHEANCE'));
TobActionGenerique.Detail[IndexAction-2].putValue('RAG_COUTACTION', Tobtampon.getvalue('RAG_COUTACTION'));
TobActionGenerique.Detail[IndexAction-2].putValue('RAG_TABLELIBRE1', Tobtampon.getvalue('RAG_TABLELIBRE1'));
TobActionGenerique.Detail[IndexAction-2].putValue('RAG_TABLELIBRE2', Tobtampon.getvalue('RAG_TABLELIBRE2'));
TobActionGenerique.Detail[IndexAction-2].putValue('RAG_TABLELIBRE3', Tobtampon.getvalue('RAG_TABLELIBRE3'));

for Nbaction := 0 to TobActionGenerique.Detail.Count-1 do
   TobActionGenerique.PutGridDetail(GAction,False,False,LesColonnes,True);
GAction.GotoRow(IndexAction-1);


inherited;

Tobtampon.free;
end;



{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 22/06/2006
Modifié le ... :   /  /
Description .. : Bouton faire descendre l'action générique
Mots clefs ... :
*****************************************************************}
procedure TFAssistCiblageVersOperation.bDescendreClick(Sender: TObject);
var
   Tobtampon : TOB;
   Nbaction : integer;
   IndexAction : integer;

begin
If TobActionGenerique.Detail.Count < 2 then     // aucune action ou juste une action
   exit;

if GAction.row = GAction.RowCount-1 then    // C'est la dernière action générique
   exit;

IndexAction := GAction.row;

//on mémorise l'action à descendre
Tobtampon := Tob.Create('ACTIONSGENERIQUES', nil, -1);
Tobtampon.putValue('RAG_TYPEACTION',TobActionGenerique.Detail[IndexAction-1].getvalue('RAG_TYPEACTION'));
Tobtampon.putValue('RAG_LIBELLE',TobActionGenerique.Detail[IndexAction-1].getvalue('RAG_LIBELLE'));
Tobtampon.putValue('RAG_DATEACTION',TobActionGenerique.Detail[IndexAction-1].getvalue('RAG_DATEACTION'));
Tobtampon.putValue('RAG_ETATACTION',TobActionGenerique.Detail[IndexAction-1].getvalue('RAG_ETATACTION'));
Tobtampon.putValue('RAG_DATEECHEANCE',TobActionGenerique.Detail[IndexAction-1].getvalue('RAG_DATEECHEANCE'));
Tobtampon.putValue('RAG_COUTACTION',TobActionGenerique.Detail[IndexAction-1].getvalue('RAG_COUTACTION'));
Tobtampon.putValue('RAG_TABLELIBRE1',TobActionGenerique.Detail[IndexAction-1].getvalue('RAG_TABLELIBRE1'));
Tobtampon.putValue('RAG_TABLELIBRE2',TobActionGenerique.Detail[IndexAction-1].getvalue('RAG_TABLELIBRE2'));
Tobtampon.putValue('RAG_TABLELIBRE3',TobActionGenerique.Detail[IndexAction-1].getvalue('RAG_TABLELIBRE3'));

//on remonte celle du dessous
TobActionGenerique.Detail[IndexAction-1].putValue('RAG_TYPEACTION', TobActionGenerique.Detail[IndexAction].getvalue('RAG_TYPEACTION'));
TobActionGenerique.Detail[IndexAction-1].putValue('RAG_LIBELLE', TobActionGenerique.Detail[IndexAction].getvalue('RAG_LIBELLE'));
TobActionGenerique.Detail[IndexAction-1].putValue('RAG_DATEACTION', TobActionGenerique.Detail[IndexAction].getvalue('RAG_DATEACTION'));
TobActionGenerique.Detail[IndexAction-1].putValue('RAG_ETATACTION', TobActionGenerique.Detail[IndexAction].getvalue('RAG_ETATACTION'));
TobActionGenerique.Detail[IndexAction-1].putValue('RAG_DATEECHEANCE', TobActionGenerique.Detail[IndexAction].getvalue('RAG_DATEECHEANCE'));
TobActionGenerique.Detail[IndexAction-1].putValue('RAG_COUTACTION', TobActionGenerique.Detail[IndexAction].getvalue('RAG_COUTACTION'));
TobActionGenerique.Detail[IndexAction-1].putValue('RAG_TABLELIBRE1', TobActionGenerique.Detail[IndexAction].getvalue('RAG_TABLELIBRE1'));
TobActionGenerique.Detail[IndexAction-1].putValue('RAG_TABLELIBRE2', TobActionGenerique.Detail[IndexAction].getvalue('RAG_TABLELIBRE2'));
TobActionGenerique.Detail[IndexAction-1].putValue('RAG_TABLELIBRE3', TobActionGenerique.Detail[IndexAction].getvalue('RAG_TABLELIBRE3'));

//on affecte le tampon à l'action en dessous
TobActionGenerique.Detail[IndexAction].putValue('RAG_TYPEACTION', Tobtampon.getvalue('RAG_TYPEACTION'));
TobActionGenerique.Detail[IndexAction].putValue('RAG_LIBELLE', Tobtampon.getvalue('RAG_LIBELLE'));
TobActionGenerique.Detail[IndexAction].putValue('RAG_DATEACTION', Tobtampon.getvalue('RAG_DATEACTION'));
TobActionGenerique.Detail[IndexAction].putValue('RAG_ETATACTION', Tobtampon.getvalue('RAG_ETATACTION'));
TobActionGenerique.Detail[IndexAction].putValue('RAG_DATEECHEANCE', Tobtampon.getvalue('RAG_DATEECHEANCE'));
TobActionGenerique.Detail[IndexAction].putValue('RAG_COUTACTION', Tobtampon.getvalue('RAG_COUTACTION'));
TobActionGenerique.Detail[IndexAction].putValue('RAG_TABLELIBRE1', Tobtampon.getvalue('RAG_TABLELIBRE1'));
TobActionGenerique.Detail[IndexAction].putValue('RAG_TABLELIBRE2', Tobtampon.getvalue('RAG_TABLELIBRE2'));
TobActionGenerique.Detail[IndexAction].putValue('RAG_TABLELIBRE3', Tobtampon.getvalue('RAG_TABLELIBRE3'));

for Nbaction := 0 to TobActionGenerique.Detail.Count-1 do
   TobActionGenerique.PutGridDetail(GAction,False,False,LesColonnes,True);
GAction.GotoRow(IndexAction+1);


inherited;

Tobtampon.free;


end;



procedure TFAssistCiblageVersOperation.MODELE_OPERATIONElipsisClick(Sender: TObject);
var
   Nbaction : integer;

begin
  inherited;
   //appel d'un lookup pour le choix de l'opération.
   //Si vrai, alors on charge les infos
  if LookupList(MODELE_OPERATION, 'Liste des Opérations', 'OPERATIONS', 'ROP_OPERATION', 'ROP_LIBELLE', '', 'ROP_OPERATION', True, -1) then
  begin
    // on efface toutes les actions génériques en cours
    for Nbaction := TobActionGenerique.Detail.Count-1  downto 0 do
    begin
      TobActionGenerique.Detail[NbAction].Free;
      GAction.DeleteRow(Nbaction+1);
    end;
    RAG_TYPEACTION.Value := '';
    RAG_LIBELLE.Text := '';
    RAG_DATEACTION.Text := '';
    RAG_ETATACTION.Value := '';
    RAG_DATEECHEANCE.Text := '';
    RAG_COUTACTION.Text := '';

    //on charge les actions génériques de l'opération choisie
    TobActionGenerique.LoadDetailDB('ACTIONSGENERIQUES','"'+MODELE_OPERATION.Text+'"','RAG_NUMACTGEN',nil,false);
    for nbaction := 0 to TobActionGenerique.Detail.count-1 do
      TobActionGenerique.PutValue('RAG_OPERATION', TobOperation.GetValue('ROP_OPERATION'));

    RechActionGenerique();

    if TobActionGenerique.Detail.Count > 0 then
    begin
      RAG_TYPEACTION.Enabled := True;
      RAG_LIBELLE.Enabled := True;
      RAG_DATEACTION.Enabled := True;
      RAG_ETATACTION.Enabled := True;
      RAG_DATEECHEANCE.Enabled := True;
      RAG_COUTACTION.Enabled := True;
      RAG_TABLELIBRE1.Enabled := True;
      RAG_TABLELIBRE2.Enabled := True;
      RAG_TABLELIBRE3.Enabled := True;


      BSupprime.Enabled := True;
      Bajoute.Enabled := True;
      BValider.Enabled := True;
      BAnnule.Enabled := True;

    end;

   end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 13/06/2007
Modifié le ... :   /  /    
Description .. : Recherche de lien sur les tablettes en correspondance 
Suite ........ : entre les ciblage et les opérations.
Mots clefs ... : 
*****************************************************************}
procedure TFAssistCiblageVersOperation.CorrespCiblageOperation();
var
  TobYLienT             : Tob;
  StrSql                : String;
  i                     : integer;
  j                     : integer;
  TobR                  : Tob;
  TobC                  : Tob;
  StrOperation          : String;
  StrCiblage            : String;
  StrRessource          : String;
  bCorr                 : Boolean;

begin
  TobYLienT             := Tob.Create('YLIENTABLETTE', nil, -1);
  StrSql                := 'SELECT * FROM YLIENTABLETTE';
  TobYLienT.LoadDetailFromSQL(StrSql);

  for i := 0 to TobYLienT.Detail.Count -1 do
  begin
    StrOperation        := '';
    StrCiblage          := '';
    StrRessource        := '';
    bCorr               := False;
    TobR                := TobYLienT.Detail[i];
    if POS('RTRROPTABLELIBRE', TobR.GetValue('YLT_DESTINATIONTAB')) > 0 then        //Tablette liée à une Opération
    begin
      StrOperation      := TobR.GetValue('YLT_DESTINATIONTAB');
      StrRessource      := TobR.GetValue('YLT_ORIGINETAB');
      if POS('AFTLIBRERES', StrRessource) > 0 then               //liée avec une ressource
      begin
        //on recherche si la ressource est liée avec un ciblage
        for j := 0 to TobYLienT.Detail.Count -1 do
        begin
          TobC          := TobYLienT.Detail[j];
          if (TobC.GetValue('YLT_ORIGINETAB') = StrRessource) and (POS('RTRCBTABLELIBRE', TobC.GetValue('YLT_DESTINATIONTAB')) > 0) then
          begin
            StrCiblage  := TobC.GetValue('YLT_DESTINATIONTAB');
            bCorr       := True;
            Break;
          end;
        end;
      end;
    end;
    if bCorr then                 //si une correspondance ciblage <-> opération
    begin
      TobR              := Tob.Create('le lien', TobLienCorr, -1);
      TobR.AddChampSupValeur('RESSOURCE', StrRessource);
      TobR.AddChampSupValeur('OPERATION', StrOperation);
      TobR.AddChampSupValeur('CIBLAGE', StrCiblage);
      TobR.AddChampSupValeur('CHOPERATION', 'ROP_' + Copy(StrOperation, 4, length(StrOperation)));
      TobR.AddChampSupValeur('CHCIBLAGE', 'RCB_' + Copy(StrCiblage, 3, length(StrCiblage)));
    end;

  end;



  FreeAndNil(TobYLienT);
end;

// FQ 10888 - STR 21/10/2008
procedure TFAssistCiblageVersOperation.RAG_TYPEACTIONExit(Sender: TObject);
var TobTypAct : TOB;
begin
  inherited;
  if RAG_TYPEACTION.value = '' then Exit;
//mcd conseil   TobTypAct := Nil;
  VH_RT.TobTypesAction.Load;
  TobTypAct:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[RAG_TYPEACTION.value,'---',0],TRUE) ;
  if TobTypAct = Nil then Exit;
  if (TobTypAct.GetValue('RPA_GESTDATECH') = 'X') then
      begin
      RAG_DATEECHEANCE.enabled := TRUE ;
      if TobTypAct.GetValue('RPA_DELAIDATECH') <> 0 then
          RAG_DATEECHEANCE.Text :=  DateToStr(RTCalculEch(StrToDate(RAG_DATEACTION.Text),StrToInt(TobTypAct.GetValue('RPA_DELAIDATECH')),TobTypAct.GetValue('RPA_WEEKEND')));
      end
  else
      begin
      RAG_DATEECHEANCE.Enabled := FALSE ;
      RAG_DATEECHEANCE.Text := ROP_DATEFIN.Text;
      end;
   if (TobTypAct.GetValue('RPA_GESTCOUT') = 'X') then
      begin
      RAG_COUTACTION.Enabled := TRUE ;
      end
   else
      begin
      RAG_COUTACTION.Enabled := FALSE ;
      RAG_COUTACTION.Text := '';
      end;
  RAG_TABLELIBRE1.value := TobTypAct.GetValue('RPA_TABLELIBRE1');
  RAG_TABLELIBRE2.value := TobTypAct.GetValue('RPA_TABLELIBRE2');
  RAG_TABLELIBRE3.value := TobTypAct.GetValue('RPA_TABLELIBRE3');
  if (TobTypAct.GetValue('RPA_REPLICLIB') = 'X' ) then
    RAG_LIBELLE.Text := RAG_TYPEACTION.Text;
  if (TobTypAct.GetString('RPA_ETATACTION') <> '' ) then
    RAG_ETATACTION.Value := TobTypAct.GetString('RPA_ETATACTION');
end;

end.
