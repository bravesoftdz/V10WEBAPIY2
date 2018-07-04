{***********UNITE*************************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 04/07/2003
Modifié le ... : 04/07/2003
Description .. : CA - 05/07/1999 - Si il n'y a plus d'amortissement, on
Suite ........ : conserve l'immo
Suite ........ : CA - 05/07/1999 - Passage à l'état "FER" des immos
Suite ........ : cédées totalement
Suite ........ : FQ 12437 - CA - 04/07/2003 - Mémorisation de la base
Suite ........ : taxe pro. pour édition à une date quelconque
Suite ........ : MBO - 27/10/2005 - gestion crc 2002/10 - depreciation d'actif
Suite ........ : MBO - 03/11/2005 - gestion crc 2002/10 - calcul et stockage de la reprise maxi de depreciation
Suite..........: FQ 17215 - TGA 21/12/2005 - GetParamSoc => GetParamSocSecur
Suite......... : BTY - 01/06 FQ 17259 Nouveau top dépréciation dans IMMO
Suite......... : BTY - 04/06 FQ 17516 Nouveau top Changement de regroupement
Suite......... : MBO - 23/05/2006 - FQ 13255 ajout des dates de l'exercice que l'on clôture
Suite......... : BTY - 05/06 FQ 18211 Positionner les indicateurs de modif de compta dans PARAMSOC
Suite......... : XVI 28/09/2006 FQ (BL) 13303 intedire la clôture si la compta n'est pas clôturé
Suite......... : MVG FQ 19236, 19237, 19244
Suite......... : MVG 05/12/06 Suite FQ 19244
Suite......... : MBO 18/04/2007 - stockage du code immo remplacée ds immolog si remplacement de composant
                                ajout de ce paramètre ds appel fonction enregistreOperationCloture
Suite......... : MBO 21/06/2007 - suite plantage en cwas : remplacement de initopeencours par le détail
Suite......... : MBO 05/07/2007 - en cwas : pas d'enreg CLO ds immolog (remplacement query par tob)
*****************************************************************}
unit imoclo;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, HPanel, HTB97, HSysMenu, hmsgbox,UiUtil, Hctrls,Hent1,
  ImEnt,HStatus, Buttons, Mask, ParamDat, ParamSoc ,utob,ImOuPlan, ImPlan
  {$IFDEF EAGLCLIENT}
  {$ELSE}
  {$IFNDEF DBXPRESS},dbtables{$ELSE},uDbxDataSet{$ENDIF}
  {$ENDIF EAGLCLIENT}
;

type
  TImParamClo = class
    bIntegOpeHisto : boolean;
    bSuppOpeHisto : boolean;
    DateSuppOpeHisto : TDateTime;
    bSuppSimul : boolean;
    bIntegFicheHisto : boolean;
    bSuppFicheHisto : boolean;
    DateSuppFicheHisto : TDateTime;
    bdelete : boolean;
  end;
  TFCloture = class(TForm)
    HPanel1: THPanel;
    GroupBox1: TGroupBox;
    cbIntegOpeHisto: TCheckBox;
    GroupBox2: TGroupBox;
    cbSuppSimul: TCheckBox;
    cbSuppOpeHisto: TCheckBox;
    DateSuppOpeHisto: THCritMaskEdit;
    cbIntegFicheHisto: TCheckBox;
    cbSuppFicheHisto: TCheckBox;
    DateSuppFicheHisto: THCritMaskEdit;
    HM: THMsgBox;
    HMTrad: THSystemMenu;
    Dock971: TDock97;
    ToolWindow971: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    dateexo: TLabel;
    debexo: TLabel;
    au: TLabel;
    finexo: TLabel;

    procedure BValiderClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure cbSuppOpeHistoClick(Sender: TObject);
    procedure cbSuppFicheHistoClick(Sender: TObject);
    procedure DateSuppOpeHistoKeyPress(Sender: TObject; var Key: Char);
    procedure DateSuppFicheHistoKeyPress(Sender: TObject; var Key: Char);
    procedure HelpBtnClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  public
    fParamClo : TImParamClo;
    function ExecuteClotureImmo: boolean ;
  private
    { Déclarations privées }
    procedure ChargeParamCloture;
    procedure MajTableImmoPourCloture ;
    procedure SetControlEnabled ( bEnabled : boolean );

  public
    { Déclarations publiques }
  end;

procedure AfficheClotureImmo; // Appel avec fiche
function GetDateMoinsDeuxAns ( dt : TDateTime) : TDateTime;
function FermetureFicheImmo (Q : TQuery) : boolean;
function SupprimeSimulation (Q  : TQuery; ParamClo : TImParamClo) : boolean;
function SupprimeHistoriqueFicheEnCours ( Q : TQuery; ParamClo : TImParamClo) : boolean;
procedure MajHistoriqueOperation ( Code : String; ParamClo : TImParamClo);
procedure SupprimeHistoriqueAncienneFiche ( Q : TQuery; ParamClo : TImParamClo);
function GetDateDerniereOperation (Code : string) : TDateTime;
procedure UpdateChampsPourCloture (var Q:TQuery);
//mbo 18/04/2007 ajout du paramètre remplace
procedure EnregistreOperationCloture (CodeImmo : string; PlanActif : integer; BaseTaxePro : double; Depreciation : double; remplace:string);

implementation

uses
  Outils
  {$IFDEF SERIE1}
  ,UtEx_, S1Util
  {$ELSE},
  UtilPGI,
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  Ent1
  {$ENDIF};

{$R *.DFM}

procedure AfficheClotureImmo ;
var FCloture: TFCloture;
    PP:THPanel;
begin
   //XVI 28/09/2006 FQ 13303 début
  {$IFDEF SERIE1}
  if not EX_IsCloturer(VHImmo^.Encours.Code) then
  begin
     PGIInfo('La clôture des immobilisations ne peut se faire que sur des exercices comptables déjà clôturés.') ;
     exit ;
  end ;
  {$ELSE}
  {$ENDIF !SERIE1}
  //XVI 28/09/2006 FQ 13303 début
  if not imBlocageMonoPoste(False) then Exit ;
  FCloture:=TFCloture.Create(Application) ;
  PP:=FindInsidePanel;
  {$IFDEF SERIE1}
  if (PP=nil) then
  {$ELSE}
  if (PP=nil) or (True) then
  {$ENDIF}
  begin
    try
      FCloture.ShowModal ;
    finally
      if FCloture.fParamClo <> nil then
      FCloture.Free ;
      ImDeblocageMonoPoste(False) ;
    end ;
  end else
  begin
    InitInside(FCloture,PP);
    FCloture.Show;
  end;
end;

function TFCloture.ExecuteClotureImmo: boolean ;
begin
result:=true ;
try
  MajTableImmoPourCloture ;
  SetParamSoc('SO_DATECLOTUREIMMO',Date) ;
  SetParamSoc('SO_EXOCLOIMMO',VHImmo^.Encours.Code) ;
  // 05/06 FQ 18211 Positionner les indicateurs de modif
  VHImmo^.ChargeOBImmo := True;
  ImMarquerPublifi(True);
  {$IFDEF SERIE1}
  if not vs1.OKModCompta then result:=EX_ClotureADistance(VHImmo^.EnCours.Code,false) ;
  {$ELSE}
  {$ENDIF}
except
  result:=false ;
end ;
end;

procedure TFCloture.FormShow(Sender: TObject);
begin
  cbIntegOpeHisto.Checked := True;
  cbSuppOpeHisto.Checked := False;
  DateSuppOpeHisto.Text := DateToStr(GetDateMoinsDeuxAns ( VHImmo^.Encours.Deb));
  cbSuppSimul.Checked := True;
  cbIntegFicheHisto.Checked := True;
  cbSuppFicheHisto.Checked := False;;
  DateSuppFicheHisto.Text := DateToStr(GetDateMoinsDeuxAns ( VHImmo^.Encours.Deb));
  DateSuppOpeHisto.Enabled := false;
  DateSuppFicheHisto.Enabled := false;
  // ajout mbo fq 13255
  Debexo.caption := DateToStr(VHImmo^.Encours.Deb);
  Finexo.caption := DateToStr(VHImmo^.Encours.Fin);
end;

procedure TFCloture.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if Parent is THPanel then
  begin
    ImDeblocageMonoPoste(False) ;
    Action:=caFree ;
  end ;
//  fParamClo.free ; YCP fait dans le destroy
end;

procedure TFCloture.BValiderClick(Sender: TObject);
var ExoClo : TExoDate;
begin
  // CA - 14/05/2002 - corrige problème des déclôture successives
  if ImQuelDateDeExo(GetParamSocSecur('SO_EXOCLOIMMO',''),ExoClo) then
  begin
    if VHImmo^.Encours.Deb <= ExoClo.Deb then
    begin
      HM.Execute(3,Caption,'');
      ModalResult := mrYes;
      exit;
    end;
  end;
{  if GetParamSoc('SO_EXOCLOIMMO')=VHImmo^.Encours.Code then
  begin
    HM.Execute(3,Caption,'');
    ModalResult := mrYes;
    exit;
  end;}
  if VHImmo^.Suivant.Code='' then
  begin
    HM.Execute(4,Caption,'');
    ModalResult := mrYes;
    exit;
  end;
  ChargeParamCloture;
  if fParamClo.DateSuppOpeHisto >= VHImmo^.EnCours.Deb then
  begin
    HM.Execute(1,Caption,'');
    FocusControl(DateSuppOpeHisto);
    ModalResult := mrNone;
  end else
  if fParamClo.DateSuppFicheHisto >= VHImmo^.EnCours.Deb then
  begin
    HM.Execute(1,Caption,'');
    FocusControl(DateSuppFicheHisto);
    ModalResult := mrNone;
  end else
  if (PGIAsk('Confirmez-vous la clôture des immobilisations ?',Caption)=mrYes) then
  begin
    SetControlEnabled ( False );
    if Transactions(MajTableImmoPourCloture,3) <> oeOK then
    begin
      HM.Execute (0,Caption,'');
      ModalResult := mrNone;
      SetControlEnabled ( True );
    end else
    BEGIN
    SetParamSoc('SO_DATECLOTUREIMMO',Date) ;
    SetParamSoc('SO_EXOCLOIMMO',VHImmo^.Encours.Code) ;
    // 05/06 FQ 18211 Positionner les indicateurs de modif
    VHImmo^.ChargeOBImmo := True;
    ImMarquerPublifi(True);
    {$IFDEF SERIE1}
    if not vs1.OKModCompta then
      begin
      if EX_ClotureADistance(VHImmo^.EnCours.Code,false) then
        HM.Execute(5,Caption,'')
        else HM.Execute(6,Caption,'') ;
      end else
      begin
      HM.Execute(2,Caption,'') ;
      end ;
    {$ELSE}
    HM.Execute(2,Caption,'') ;
    {$ENDIF}
    ModalResult := MrYes;
    END ;
  end else
  begin
    SetControlEnabled ( True );
    ModalResult := mrNone;
  end;
end;

procedure TFCloture.BFermeClick(Sender: TObject);
begin
  ModalResult := mrCancel;
end;

procedure TFCloture.cbSuppOpeHistoClick(Sender: TObject);
begin
  DateSuppOpeHisto.enabled := cbSuppOpeHisto.Checked;
end;

procedure TFCloture.cbSuppFicheHistoClick(Sender: TObject);
begin
  DateSuppFicheHisto.enabled := cbSuppFicheHisto.Checked;
end;

procedure TFCloture.DateSuppOpeHistoKeyPress(Sender: TObject;
  var Key: Char);
begin
  ParamDate (Self, Sender, Key);
end;

procedure TFCloture.DateSuppFicheHistoKeyPress(Sender: TObject;
  var Key: Char);
begin
  ParamDate (Self, Sender, Key);
end;

procedure TFCloture.ChargeParamCloture;
begin
  fParamClo.bIntegOpeHisto := cbIntegOpeHisto.Checked;
  fParamClo.bSuppOpeHisto := cbSuppOpeHisto.Checked;
  fParamClo.DateSuppOpeHisto := StrToDate(DateSuppOpeHisto.Text);
  fParamClo.bSuppSimul := cbSuppSimul.Checked;
  fParamClo.bIntegFicheHisto := cbIntegFicheHisto.Checked;
  fParamClo.bSuppFicheHisto := cbSuppFicheHisto.Checked;
  fParamClo.DateSuppFicheHisto := StrToDate(DateSuppFicheHisto.Text);
end;

function GetDateMoinsDeuxAns ( dt : TDateTime) : TDateTime;
begin
  result := PlusMois(dt,-24);
end;

procedure TFCloture.MajTableImmoPourCloture ;
var Q : TQuery; OkPost : boolean; stRequete: string;
begin
{ CA - 07/05/2003 - On ne prend pas en compte les immobilisations de N+1 }
stRequete := 'SELECT * FROM IMMO WHERE I_DATEPIECEA<="'+USDateTime(VHImmo^.EnCours.Fin)+'"';
Q:=OpenSQL(stRequete,False);
InitMove(100,'');
while not Q.Eof do
  begin
  MoveCur(FALSE) ;
  fparamclo.bdelete:=false;
  if Q.FindField('I_ETAT').AsString='FER' then
    SupprimeHistoriqueAncienneFiche (Q,fParamClo)
  else
    if not SupprimeSimulation(Q,fParamClo) then
      begin
      OkPost:=not FermetureFicheImmo(Q);
      if not OkPost then OkPost:=not SupprimeHistoriqueFicheEnCours(Q,fParamClo);
      if OkPost then Q.Post;
      MajHistoriqueOperation(Q.FindField('I_IMMO').AsString ,fParamClo);
      end;
    if fparamclo.bdelete=false then q.next; //FQ 19236 et 19237 si q.delete ne pas faire de q.next sinon ca saute l'immo suivante
  end;
FiniMove;
Ferme (Q);
end;

function FermetureFicheImmo (Q : TQuery) : boolean;
var Etat, Nature : string;
    Plan : TPlanAmort;
    depreciation : double;
    remplace : string;

begin
  // Si la date de dernier mvt eco et fiscal est antérieure à la date de
  // fin d'exercice, on ferme la fiche.
  Result := True;
  Etat := Q.FindField('I_ETAT').AsString;
  Nature := Q.FindField('I_NATUREIMMO').AsString;
  Remplace := Q.FindField('I_REMPLACE').AsString;

  if  Etat = 'FER' then exit;
  Result := false;
  if (Q.FindField('I_DATEPIECEA').AsDateTime <= VHImmo^.EnCours.Fin) then
  begin
    if ((Q.FindField('I_QUANTITE').AsInteger=0) and (Q.FindField('I_OPECESSION').AsString='X')) then // Immo cédée totalement
    begin
      Q.Edit;
      Q.FindField('I_ETAT').AsString := 'FER';
      // BTY 01/06 FQ 17259 Nouveau top dépréciation
      // BTY 04/06 FQ 17516
      // modif mbo erreur en cwas - 21.06.07
      //InitOpeEnCoursImmo(Q,'-','-','-','-','-','-','-','-','-','-','-');
      Q.FindField('I_OPEMUTATION').AsString := '-';
      Q.FindField('I_OPEECLATEMENT').AsString := '-';
      Q.FindField('I_OPECESSION').AsString := '-';
      Q.FindField('I_OPECHANGEPLAN').AsString := '-';
      Q.FindField('I_OPELIEUGEO').AsString := '-';
      Q.FindField('I_OPEETABLISSEMENT').AsString := '-';
      Q.FindField('I_OPELEVEEOPTION').AsString := '-';
      Q.FindField('I_OPEMODIFBASES').AsString := '-';
      Q.FindField('I_OPEDEPREC').AsString := '-';
      Q.FindField('I_OPEREG').AsString := '-';
      Q.FindField('I_OPERATION').AsString := '-';

      Result := True;
    end;
    if ((Nature = 'CB') or (Nature = 'LOC')) and
       ((Q.FindField ('I_DATESUSPCB').AsDateTime >= Q.FindField ('I_DATEPIECEA').AsDateTime)
           and (Q.FindField ('I_DATESUSPCB').AsDateTime <= VHImmo^.Encours.Fin))
           or  (Q.FindField('I_OPELEVEEOPTION').AsString='X') then
    begin
      Q.Edit;
      Q.FindField('I_ETAT').AsString := 'FER';
      // BTY 01/06 FQ 17259 Nouveau top dépréciation
      // BTY 04/06 FQ 17516
      // modif mbo erreur en cwas - 21.06.07
      //InitOpeEnCoursImmo(Q,'-','-','-','-','-','-','-','-','-','-','-');
      Q.FindField('I_OPEMUTATION').AsString := '-';
      Q.FindField('I_OPEECLATEMENT').AsString := '-';
      Q.FindField('I_OPECESSION').AsString := '-';
      Q.FindField('I_OPECHANGEPLAN').AsString := '-';
      Q.FindField('I_OPELIEUGEO').AsString := '-';
      Q.FindField('I_OPEETABLISSEMENT').AsString := '-';
      Q.FindField('I_OPELEVEEOPTION').AsString := '-';
      Q.FindField('I_OPEMODIFBASES').AsString := '-';
      Q.FindField('I_OPEDEPREC').AsString := '-';
      Q.FindField('I_OPEREG').AsString := '-';
      Q.FindField('I_OPERATION').AsString := '-';

      Result := True;
    end;
  end;
  if (Etat <> 'FER') then
  begin
    if ((Etat = 'OUV') and (Q.FindField('I_DATEPIECEA').AsDateTime <= VHImmo^.EnCours.Fin)) then
    begin
      // Si la fiche est ouverte, l'état passe à CLO pour indiquer que la fiche
      // a déjà été concernée par une clôture
      if not Result then
      begin
        Q.Edit;
        UpdateChampsPourCloture (Q);
        if (Q.FindField('I_DATEPIECEA').AsDateTime<=VHImmo^.Encours.Fin) then
          Q.FindField('I_ETAT').AsString := 'CLO';
      end;
    end else if (Etat='CLO') then
    begin
      Q.Edit;
      UpdateChampsPourCloture (Q);
    end;
    // Mise à jour du plan d'amortissement
    if (Q.FindField('I_DATEPIECEA').AsDateTime <= VHImmo^.EnCours.Fin) then
    begin { Pour les immobilisations en cours }
      Plan:=TPlanAmort.Create(true) ; // := CreePlan(True);
      try
        Plan.Charge(Q);
        Plan.Recupere(Q.FindField('I_IMMO').AsString,Q.FindField('I_PLANACTIF').AsString);
        // ajout mbo 3.11.05
        Depreciation := Plan.CalculRepriseDepreciation(Plan.Amorteco,VHImmo^.Suivant.Deb,
                                                  VHImmo^.Suivant.Fin, true);
        Q.FindField('I_REVISIONECO').AsFloat := 0;

        Plan.CalculPourTraitement ('CLO', Depreciation);
        Plan.Sauve;
        Q.FindField('I_PLANACTIF').AsInteger := Plan.NumSeq;
      finally
        Plan.free ; //Detruit;
      end ;
      EnregistreOperationCloture (Q.FindField('I_IMMO').AsString,Q.FindField('I_PLANACTIF').AsInteger, Q.FindField('I_BASETAXEPRO').AsFloat, Depreciation, remplace);
    end else
    begin { Pour les immobilisations sur des exercices postérieurs à l'en-cours }
      Plan:=TPlanAmort.Create(true) ; // := CreePlan(True);
      try
        Q.Edit;
        Plan.Charge(Q);
        Plan.Recupere(Q.FindField('I_IMMO').AsString,Q.FindField('I_PLANACTIF').AsString);
        // modif mbo 4.11.05 Plan.CalculPourTraitement ('CLO');
        Plan.CalculPourTraitement ('CLO',0);
        { Suppression du plan d'amortissement }
        ExecuteSQL ('DELETE FROM IMMOAMOR WHERE IA_IMMO="'+Q.FindField('I_IMMO').AsString+'"');
        Plan.NumSeq := Plan.NumSeq-1;
        Plan.Sauve;
      finally
        Plan.free ; //Detruit;
      end ;
    end;
  end;
end;

// nouveau pour crc2002-10
{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 03/11/2005
Modifié le ... :   /  /
Description .. : permet de calculer la reprise maximale possible quand il a
Suite ........ : été saisie une dépréciation d'actif.
Mots clefs ... :
*****************************************************************}
{ méthodes de calcul utilisées avec N = exercice que l'on cloture et N+1 = exercice que l'on ouvre

  - mode linéaire =
     reprise maxi fin exercice N * ( 1- Durée exercice N+1 / durée restante debut N+1)

  - mode variable =
     reprise maxi fin exercice N * ( 1 - nb UO exercice N+1 / nb UO restant debut exo N+1)

  mode dégressif =
     reprise maxi fin exercice N * (1 - taux )
}
{function CalculRepriseMaxi(Plan:TPlanAmort; Codeimmo:String): double;
Var QLog: TQuery;
    type_op : string;
    Cumul : Double;
    nb1 : Integer;
    nb2 : Integer;
    Prorata:Double;
    Montant:Double;
    TauxCalcule:Double;
begin
  Cumul :=0.00;
  Montant:=0.00;

  // impact antérieur dépréciation - mbo 02.06
  if Plan.amortEco.RepriseDep <> 0 then
     Plan.CalculRepriseDepreciation(fplan.AmortEco, VHImmo^.Suivant.Deb,
                                               VHImmo^.Suivant.Fin, false);

  type_op:='CLO';

  QLog := OpenSQL('SELECT IL_REVISIONECO FROM IMMOLOG WHERE (IL_IMMO="'
      + CodeImmo + '" AND IL_TYPEOP ="' + type_op +'" AND IL_DATEOP<"' +
      USDateTime(VHImmo^.Suivant.Deb) +'")' , True);

  while not QLog.Eof do
  begin
      Cumul :=  QLog.FindField('IL_REVISIONECO').AsFloat;
      QLog.Next;
  end;
  Ferme(QLog);

  if (Cumul = 0) and (CumulAnt <> 0) then
     Cumul := CumulAnt;
     Cumul := Cumul + Plan.RevisionEco;
end;

  if Plan.AmortEco.Methode = 'LIN' then
  begin
    //nb1 = nb jours de l'exercice - nb2 = nb jours restant à amortir
    nb1 := NOMBREJOUR360 (VHImmo^.Suivant.Deb,VHImmo^.Suivant.Fin);
    nb2 := NOMBREJOUR360 (VHImmo^.Suivant.Deb, Plan.AmortEco.DateFinAmort);

    if nb1 > nb2 then   // cas de la dernière année
      montant := 0.00
    else
      montant := cumul * (1 - (nb1/nb2));

  end else
    if Plan.AmortEco.Methode = 'VAR' then
    begin
       prorata:= Plan.ProrataUOreste(VHImmo^.Encours.Fin);
       montant:= cumul * (1-prorata);
    end else

    if Plan.AmortEco.methode = 'DEG' then
    begin
       TauxCalcule := Plan.DegressifTauxReprise(Plan.AmortEco,VHImmo^.Suivant.Deb, true);
       montant := cumul * (1-TauxCalcule);
  end;

  Result := arrondi(montant,V_PGI.OkDecV);

end; }

procedure EnregistreOperationCloture (CodeImmo : string; PlanActif : integer; BaseTaxePro : double; Depreciation : double; remplace:string);
var Ordre,OrdreS : integer;
    //QUeryW : TQuery;
    QTobLog : TOB;
begin
  Ordre := TrouveNumeroOrdreLogSuivant(CodeImmo);
  OrdreS := TrouveNumeroOrdreSerieLogSuivant;

  // mbo le 5.07.07 remplacé par une tob car ne fonctionnait pas en cwas
  {QueryW:=OpenSQL('SELECT * FROM IMMOLOG WHERE IL_IMMO="'+W_W+'"', FALSE) ;
  QueryW.Insert ;
  QueryW.FindField('IL_IMMO').AsString:=CodeImmo ;
  QueryW.FindField('IL_LIBELLE').AsString:=RechDom('TIOPEAMOR', 'CLO', FALSE)+' '+DateToStr(VHImmo^.Encours.Fin+1);
  QueryW.FindField('IL_TYPEMODIF').AsString:=AffecteCommentaireOperation('CLO');
  QueryW.FindField('IL_DATEOP').AsDateTime:=VHImmo^.Encours.Fin;
  QueryW.FindField('IL_TYPEOP').AsString:='CLO' ;
  QueryW.FindField('IL_PLANACTIFAV').AsFloat:=PlanActif - 1 ;
  QueryW.FindField('IL_PLANACTIFAP').AsFloat:=PlanActif ;
  QueryW.FindField('IL_ORDRE').AsInteger:= Ordre;
  QueryW.FindField('IL_ORDRESERIE').AsInteger:=OrdreS;
  { FQ 12437 - CA - 04/07/2003 - Mémorisation de la base taxe pro. pour édition à une date quelconque }
  {QueryW.FindField('IL_BASETAXEPRO').AsFloat := BaseTaxePro;
  // ajout mbo 3.11.2005  stockage de la reprise maximale
  QueryW.FindField('IL_REVISIONECO').AsFloat := Depreciation;
  // ajout mbo 18/04/2007 stockage de l'immo remplacée si remplacement de composant
  QueryW.FindField('IL_CODEMUTATION').AsString := Remplace;
  QueryW.Post;
  Ferme(QueryW) ;  }

  QTobLog := TOB.Create('IMMOLOG',nil,-1);
  QtobLog.PutValue('IL_IMMO', CodeImmo) ;
  QtobLog.PutValue('IL_LIBELLE',RechDom('TIOPEAMOR', 'CLO', FALSE)+' '+DateToStr(VHImmo^.Encours.Fin+1));
  QtobLog.PutValue('IL_TYPEMODIF',AffecteCommentaireOperation('CLO'));
  QtobLog.PutValue('IL_DATEOP',VHImmo^.Encours.Fin);
  QtobLog.PutValue('IL_TYPEOP','CLO');
  QtobLog.PutValue('IL_PLANACTIFAV',(PlanActif - 1));
  QtobLog.PutValue('IL_PLANACTIFAP',PlanActif);
  QtobLog.PutValue('IL_ORDRE',Ordre);
  QtobLog.PutValue('IL_ORDRESERIE',OrdreS);
  { FQ 12437 - CA - 04/07/2003 - Mémorisation de la base taxe pro. pour édition à une date quelconque }
  QtobLog.PutValue('IL_BASETAXEPRO',BaseTaxePro);
  QtobLog.PutValue('IL_REVISIONECO',Depreciation);
  // ajout mbo 18/04/2007 stockage de l'immo remplacée si remplacement de composant
  QtobLog.PutValue('IL_CODEMUTATION',Remplace);
  QtobLog.InsertDB(nil,False);
  QtobLog.Free;

end;

procedure UpdateChampsPourCloture (var Q:TQuery);
begin
  Q.FindField('I_REPCEDECO').AsFloat := 0;
  Q.FindField('I_REPCEDFISC').AsFloat := 0;
  Q.FindField('I_VNC').AsFloat := 0;
  Q.FindField('I_VOACEDE').AsFloat := 0;
  Q.FindField('I_QTCEDE').AsFloat := 0;
  Q.FindField('I_MONTANTEXC').AsFloat := 0;
  Q.FindField('I_MONTANTEXCCED').AsFloat := 0;
  Q.FindField('I_TYPEEXC').AsString := '';
  Q.FindField('I_BASEAMORDEBEXO').AsFloat := Q.FindField('I_MONTANTHT').AsFloat ;

  // MBO - 18/04/2007 - opération de remplacement de composant
  Q.FindField('I_TYPER').AsString := '';
  Q.FindField('I_REMPLACE').AsString := '';

  // BTY 01/06 FQ 17259 Nouveau top dépréciation
  // BTY 04/06 FQ 17516
  // modif mbo erreur en cwas - 21.06.07
  //InitOpeEnCoursImmo(Q,'-','-','-','-','-','-','-','-','-','-','-');
  Q.FindField('I_OPEMUTATION').AsString := '-';
  Q.FindField('I_OPEECLATEMENT').AsString := '-';
  Q.FindField('I_OPECESSION').AsString := '-';
  Q.FindField('I_OPECHANGEPLAN').AsString := '-';
  Q.FindField('I_OPELIEUGEO').AsString := '-';
  Q.FindField('I_OPEETABLISSEMENT').AsString := '-';
  Q.FindField('I_OPELEVEEOPTION').AsString := '-';
  Q.FindField('I_OPEMODIFBASES').AsString := '-';
  Q.FindField('I_OPEDEPREC').AsString := '-';
  Q.FindField('I_OPEREG').AsString := '-';
  Q.FindField('I_OPERATION').AsString := '-';
end;

function SupprimeSimulation (Q  : TQuery; ParamClo : TImParamClo) : boolean;
var Code : string;
begin
  Result := False;
  if (Q.FindField('I_DATEPIECEA').AsDateTime>VHImmo^.EnCours.Fin) then exit;
  if (ParamClo.bSuppSimul) and
           (Q.FindField('I_QUALIFIMMO').AsString='S') then
  begin
    Code := Q.FindField ('I_IMMO').AsString ;
    Q.Delete;
    ExecuteSQL ('DELETE FROM IMMOLOG WHERE IL_IMMO="'+Code+'"');
    ExecuteSQL ('DELETE FROM IMMOAMOR WHERE IA_IMMO="'+Code+'"');
    ExecuteSQL ('DELETE FROM IMMOECHE WHERE IH_IMMO="'+Code+'"');
    ExecuteSQL ('DELETE FROM IMMOUO WHERE IUO_IMMO="'+Code+'"');     // MVG 05/12/06
    paramclo.bdelete:=true;
    result := True;
  end else Result := false;
end;

function SupprimeHistoriqueFicheEnCours ( Q : TQuery; ParamClo : TImParamClo) : boolean;
var Code : string;
begin
  Result := false;
  if not ParamClo.bIntegFicheHisto then
  begin
    // FQ 19244
    Code := Q.FindField ('I_IMMO').AsString ;
    Q.Delete;
    // FQ 19244
    ExecuteSQL ('DELETE FROM IMMOLOG WHERE IL_IMMO="'+Code+'"');
    ExecuteSQL ('DELETE FROM IMMOAMOR WHERE IA_IMMO="'+Code+'"');
    ExecuteSQL ('DELETE FROM IMMOECHE WHERE IH_IMMO="'+Code+'"');
    ExecuteSQL ('DELETE FROM IMMOUO WHERE IUO_IMMO="'+Code+'"');     // MVG 05/12/06
    paramclo.Bdelete:=true;
    Result := True;
  end;
end;

procedure SupprimeHistoriqueAncienneFiche (Q :TQuery; ParamClo :TImParamClo);
var Code : string;
begin
  if not ParamClo.bSuppFicheHisto then exit;
  if ((Q.FindField('I_DATESUSPCB').AsDateTime<=ParamClo.DateSuppFicheHisto) and
             (Q.FindField ('I_DATESUSPCB').AsDateTime>iDate1900))
      or (GetDateDerniereOperation (Q.FindField('I_IMMO').AsString) <= ParamClo.DateSuppFicheHisto)
  then
  begin
  // FQ 19244
  Code := Q.FindField ('I_IMMO').AsString ;
  Q.Delete;
  // FQ 19244
  ExecuteSQL ('DELETE FROM IMMOLOG WHERE IL_IMMO="'+Code+'"');
  ExecuteSQL ('DELETE FROM IMMOAMOR WHERE IA_IMMO="'+Code+'"');
  ExecuteSQL ('DELETE FROM IMMOECHE WHERE IH_IMMO="'+Code+'"');
  ExecuteSQL ('DELETE FROM IMMOUO WHERE IUO_IMMO="'+Code+'"');     // MVG 05/12/06
  paramclo.Bdelete:=true;
  end;
end;

function GetDateDerniereOperation (Code : string) : TDateTime;
var Q : TQuery;
begin
  Result := iDate1900;
  Q := OpenSQL ('SELECT IL_DATEOP FROM IMMOLOG WHERE IL_IMMO="'+Code+'" ORDER BY IL_DATEOP DESC', True);
  if not Q.Eof then
  begin
    Q.First;
    Result := Q.FindField('IL_DATEOP').AsDateTime;
  end;
  Ferme (Q);
end;

procedure MajHistoriqueOperation ( Code : String; ParamClo : TImParamClo);
var stWhere : string;
begin
  if (ParamClo.bIntegOpeHisto) and (not ParamClo.bSuppOpeHisto) then exit;
  if (not ParamClo.bIntegOpeHisto) and (not ParamClo.bSuppOpeHisto) then
    stWhere := ' AND (IL_DATEOP<="'+USDateTime(VHImmo^.Encours.Fin)+
    '" AND IL_DATEOP>="'+USDateTime(VHImmo^.Encours.Deb)+'")' else
  if (not ParamClo.bIntegOpeHisto) and (ParamClo.bSuppOpeHisto) then
    stWhere := ' AND (IL_DATEOP<="'+USDateTime(VHImmo^.Encours.Fin)+
    '" AND IL_DATEOP>="'+USDateTime(VHImmo^.Encours.Deb)+'") OR (IL_DATEOP<="'+
    USDateTime(ParamClo.DateSuppOpeHisto)+'")' else
  if (ParamClo.bIntegOpeHisto) and (ParamClo.bSuppOpeHisto) then
    stWhere := ' AND (IL_DATEOP<="'+USDateTime(ParamClo.DateSuppOpeHisto)+'")';
//  ExecuteSQL ('DELETE FROM IMMOLOG WHERE IL_IMMO="'+Code+'"'+stWhere);
  { On ne supprime pas l'enregistrement d'acquisition }
  ExecuteSQL ('DELETE FROM IMMOLOG WHERE IL_IMMO="'+Code+'" AND IL_TYPEOP<>"ACQ" '+stWhere);
end;


procedure TFCloture.HelpBtnClick(Sender: TObject);
begin
CallHelpTopic(Self);
end;

procedure TFCloture.FormCreate(Sender: TObject);
begin
fParamClo:=TImParamClo.Create ;
{$IFDEF SERIE1}
HelpContext:=541600 ;
{$ELSE}
HelpContext:=2416000 ;
{$ENDIF}
end;

//XMG 22/11/02 début
//////////////////////////////////////////////////////////////////////////////////
procedure TFCloture.FormDestroy(Sender: TObject);
begin
FParamClo.Free ;
end;
//XMG 22/11/02 fin
procedure TFCloture.SetControlEnabled(bEnabled: boolean);
begin
  cbIntegOpeHisto.Enabled := bEnabled;
  cbSuppOpeHisto.Enabled := bEnabled;
  DateSuppOpeHisto.Enabled := bEnabled;
  cbSuppSimul.Enabled := bEnabled;
  cbIntegFicheHisto.Enabled := bEnabled;
  cbSuppFicheHisto.Enabled := bEnabled;
  DateSuppFicheHisto.Enabled := bEnabled;
  BValider.Enabled := bEnabled;
  BFerme.Enabled := bEnabled;
end;

end.
