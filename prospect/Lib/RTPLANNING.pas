unit RTPLANNING;

interface

uses
  SysUtils, Classes, Graphics, Controls, Forms,
  ExtCtrls, StdCtrls, Hctrls, HTB97,Utob,Hplanning,HPanel,HEnt1,UIUtil,EntRT,
  ParamDat,HMsgBox,Paramsoc,M3FP,UtilRT,heureutil,AGLInit,
  AGLInitGC,YRessource,Yplanning,UtilAction,windows, UdateUtils,
{$IFDEF EAGLCLIENT}
 MaineAGL, Mask,
{$ELSE}
  Fe_main {$IFNDEF DBXPRESS},dbtables{BDE}{$ELSE},uDbxDataSet{$ENDIF},MailOl
{$ENDIF}
  ;

type
  RecordPlanning = Record
    TobItems : TOB ;
    TobEtats : TOB ;
    TobRes   : TOB ;
    TobCols  : TOB ;
    TobRows  : TOB ;
    TobEvents: TOB ;
    TobResAff: TOB ;
  end ;

type
  RecordCritere = Record
    StIntervenant : string;
    StTiers       : string;
    StTypeActions : string;
  end ;

  TRT_Planning = class(TForm)
    Panneau: TPanel;
    PPlanning: TPanel;
    BCalendrier: TToolbarButton97;
    DateEdit: THCritMaskEdit;
    BQuitter: TToolbarButton97;
    BPageSuiv: TToolbarButton97;
    BPagePrec: TToolbarButton97;
    Bimprimer: TToolbarButton97;
    RESSOURCE: THCritMaskEdit;
    TRessource: TLabel;
    BAide: TToolbarButton97;
    BRecharger: TToolbarButton97;
    BEnvoieOutlook: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure BCalendrierClick(Sender: TObject);
    procedure BPageSuivClick(Sender: TObject);
    procedure BPagePrecClick(Sender: TObject);
    procedure BQuitterClick(Sender: TObject);
    procedure BimprimerClick(Sender: TObject);
    function MajTobItem(Tobitem: TOB;FromItem: TOB): Boolean;
    function DroitModif(Tobitem: TOB;TypeModif: string): Boolean;
    procedure RESSOURCEElipsisClick(Sender: TObject);
    function DureeOK(Tobitem: TOB): Boolean;
    procedure BRechargerClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure RESSOURCEExit(Sender: TObject);
    procedure BEnvoieOutlookClick(Sender: TObject);
  private
    { Déclarations privées }
    TobModelePlanning: Tob;
    DateSelection,DateFin : TDateTime;
    Saisie : Integer;
    ChgtRessource : Boolean;
    NumModele : Integer;
    Criteres : RecordCritere;
    Procedure MajEvenements (Var MajPlanning:THPlanning);
    procedure Creation(Sender: TObject; Item: TOB; var Cancel: boolean);
    procedure InitItem(Sender: TObject; var Item: TOB; var Cancel: boolean);
    procedure Modification(Sender: TObject; Item: TOB; var Cancel: Boolean);
    procedure Suppression(Sender: TObject; Item: TOB; var Cancel: boolean);
    procedure DoubleClick(Sender: Tobject);
    function LoadModeles : boolean;
    function LoadTobParam(T:TOB): integer;
    Procedure ChargeEtatPlanning  ;
    function LoadTobEtat(T:TOB): integer;
    procedure AvertirApplication(Sender: TObject; FromItem: TOB; ToItem: TOB; Actions: THPlanningAction);
    Procedure P1BeforeChange(const Item: TOB;const LaRessource: String; const LaDateDeDebut, LaDateDeFin: TDateTime;var Cancel: Boolean);
    procedure Link(Sender: TObject; TobSource, TobDestination: TOB;
      Option: THPlanningOptionLink; var Cancel: Boolean);
    procedure Deplacement(Sender: TObject; Item: TOB; var Cancel: boolean);
    procedure OnPopup(Item: tob; ZCode: integer; var Redraw: boolean);
    function ControleEstLibre(Item,TobRessources:TOB): Boolean;
  public
    { Déclarations publiques }
    Planning: THPlanning;
    TobPlannings: RecordPlanning ;
  end;

Const
    TypeAgenda    = 0;
    TypePlanning  = 1;
    NbPoPup = 2;
    IndPoPupContact = 0;
const
	// libellés des messages
	TexteMessage: array[1..8] of string 	= (
          {1}        'Vous n''êtes pas autorisé à modifier cette action'
          {2}        ,'Cette action n''est pas supprimable'
          {3}        ,'Pas d''activité sur cette période'
          {4}        ,'La durée ne peut pas dépasser 12 heures'
          {5}        ,'Vous ne pouvez pas changer le type d''action par un déplacement'
          {6}        ,'La mise à jour de la base n''est pas faite'
          {7}        ,'Confirmez-vous la suppression?'
          {8}       ,'Des intervenants ne sont pas disponibles sur cette période'
          );

                                   CodeEtatAct : array[1..3] of String 	= (
                      ('PRE'),
                      ('REA'),
                      ('NRE')
          );

var
  RT_Planning: TRT_Planning;

procedure ExecutePlanning (Criteres : string; DateDeb,DateFin : TDateTime; Saisie : Integer; ChgtRessource:Boolean;TypePlanning: string);
procedure ChargeParamPlanning (var R: RecordPlanning ;var P:THPlanning ; T:TOB ; DateEnCours,DateFin:TDateTime; Criteres:RecordCritere; Mode,NumModele:integer ) ;
procedure InitParamPlanning (var P:THPlanning; T:TOB; Mode,NumModele : integer );
//function ChargeResPlanning (var R: RecordPlanning): Boolean ;
Procedure ChargeItemsPlanning (var R: RecordPlanning; T:TOB ; DateDeb,DateFin:TDateTime; Criteres: RecordCritere; NumModele: integer) ;
Procedure AffichagePlanning (var R: RecordPlanning;var P: THPlanning;DateEnCours,DateFin:TDateTime);
function ChargeResAff (var R: RecordPlanning ; Criteres: RecordCritere): Boolean ;
procedure CompleteRequete(var Requete: string; Criteres: RecordCritere);
function DateToFloat(TDate: TDateTime): Double;
procedure CalculDateRappel(ToItem:TOB ;FromItem: TOB);
procedure ReajustementDates(Tobitem: TOB;T:TOB);
function GetDescriptifCourt ( TOBD : TOB) : string;

implementation

{$R *.DFM}

uses rtfCounter,CalcOLEGenericBTP;

function CtrlHeure(Tobitem: TOB;T:TOB): Boolean;
var Hour, Min, Sec, MSec: Word;
    Dureeact : double;
begin
DecodeTime(TobItem.GetValue('RAC_HEUREACTION'), Hour, Min, Sec, MSec);
Dureeact := TobItem.getValue('RAC_DUREEACTION');
Result := True;
if ((Min <> 0) and (Min <> 30)) or (((Min = 30) or (Frac(Dureeact/60) = 0.5)) and (T.GetValue('CADENCEMENT') = '005') or (Dureeact = 0)) then Result := False;
end;

procedure ExecutePlanning (Criteres : string; DateDeb,DateFin : TDateTime ;Saisie : Integer; ChgtRessource : Boolean;TypePlanning:string);
Var PPANEL : THPanel ;
    Fo_Planning : TRT_Planning;
    Critere,ChampMul,ValMul : string;
    x : integer;
begin
  Fo_Planning := TRT_Planning.Create ( Application );
  Fo_Planning.DateSelection := DateDeb;
  Fo_Planning.DateFin := DateFin;
  Fo_Planning.Saisie := Saisie;
  Fo_Planning.ChgtRessource := ChgtRessource;
  if (TypePlanning = 'AGENDA') then
     begin
     Fo_Planning.NumModele := 0;
     Fo_Planning.HelpContext := 111000361;
     Fo_Planning.Width := 640;
     Fo_Planning.Height := 553;
     end
  else
     if (TypePlanning = 'PLANNING') then
     begin
     Fo_Planning.NumModele := 1;
     Fo_Planning.HelpContext := 111000362;
     Fo_Planning.Width := 790;
     Fo_Planning.Height := 550;
     Fo_Planning.Left := 230;
     end;
  Repeat
      Critere:=Trim(ReadTokenPipe(Criteres,'|')) ;
      if Critere<>'' then
      begin
          x:=pos('=',Critere);
          if x<>0 then
          begin
             ChampMul:=copy(Critere,1,x-1);
             ValMul:=copy(Critere,x+1,length(Critere));
             if ChampMul='INTERVENANT' then
             begin
                Fo_Planning.Criteres.StIntervenant := ValMul;
             end;
             if ChampMul='TIERS' then
             begin
                Fo_Planning.Criteres.StTiers := ValMul;
             end;
             if ChampMul='TYPEACTIONS' then
             begin
                Fo_Planning.Criteres.StTypeActions := ValMul;
             end;
          end;
      end;
  until  Critere='';
{   inside := Nil;
  Try
  Inside:=FindInsidePanel ;
   InitInside(X,Inside) ;
      X.Show ;
    finally
  end;     }
{PPANEL := FindInsidePanel ; // permet de savoir si la forme dépend d'un PANEL
if PPANEL = Nil then
  begin
      try
        Fo_Planning.ShowModal ;
      finally
        Fo_Planning.Free ;
      end ;
   end else
   BEGIN
   InitInside (Fo_Planning, PPANEL);
   Fo_Planning.Show ;
   end ;  }
{
  Try
        Fo_Planning.ShowModal ;
  finally
        Fo_Planning.Free ;
  end;       }  // fin fct ok

  {$IFDEF EAGLCLIENT}
  PPANEL:=FindInsidePanel ;
  {$ELSE}
  PPANEL:=Nil;
  {$ENDIF}
  if (PPANEL=NIL) or (V_PGI.ZoomOLE = true) then   // Fiche 10282 (l'AGENDA s'affichait derrière la fenêtre "Mes Actions")
  BEGIN
//    X.FindPrien := Nil ;
    try
      Fo_Planning.ShowModal ;
    finally
      Fo_Planning.Free ;
    End;
  end
  Else
  Begin
   InitInside(Fo_Planning,PPANEL) ;
   Fo_Planning.Show ;
   END ;
end;

procedure TRT_Planning.FormShow(Sender: TObject);
var StCaption,ListeActions : string;
    Q : TQuery ;
begin
inherited;
Planning := THPlanning.Create(Self);
Planning.parent := PPlanning;
MajEvenements (Planning);
TobPlannings.TobItems   := TOB.Create('Les_items',Nil,-1) ;
TobPlannings.TobRes     := TOB.Create('Les_ressources',Nil,-1) ;
TobPlannings.TobEtats   := TOB.Create('Les_etats',Nil,-1) ;
TobPlannings.TobResAff  := TOB.Create('Les_ressources_aff',Nil,-1) ;
TobModelePlanning   := TOB.Create('Les_modeles',Nil,-1) ;
Case NumModele of
   TypeAgenda:
   begin
   BImprimer.visible := False;
   DateEdit.Text := DateTimeToStr(DateSelection);
   Ressource.text := Criteres.StIntervenant;
   if ChgtRessource = False then Ressource.Enabled := False;
   Caption := 'Agenda de '+ RechDom ('RTREPRESENTANT1',Criteres.StIntervenant,False);
   end;
   TypePlanning:
   begin
   DateEdit.visible := False;
   Ressource.visible := False;
   TRessource.visible := False;
   BCalendrier.visible := False;
   BPagePrec.visible := False;
   BPageSuiv.visible := False;
   StCaption := 'Actions du '+DateTimeToStr(DateSelection) + ' au '+DateTimeToStr(DateFin);
   if Criteres.StIntervenant <> '' then
     StCaption := StCaption + ' de ' + RechDom ('RTREPRESENTANT1',Criteres.StIntervenant,False);
   if Criteres.StTiers <> '' then
     begin
     Q:=OpenSQL('SELECT T_LIBELLE FROM TIERS WHERE T_TIERS = "'+ Criteres.StTiers +'"', True,-1,'',true);
     if not Q.Eof then StCaption := StCaption + ' pour le client ' + Q.FindField('T_LIBELLE').asstring
     else StCaption := StCaption + ' pour le client ' + Criteres.StTiers;
     Ferme(Q) ;
     end;
   Caption := StCaption;
   end;
end;
UpdateCaption (Self);

if not LoadModeles then
  begin
//  HShowMessage('', 'Attention !', 'Le paramétrage du planning n''a pas été défini');
  end;
Case NumModele of
   TypeAgenda:
   begin
   Criteres.StTypeActions := TOBModelePlanning.detail[0].Getvalue('TYPEACTIONS');
   end;
   TypePlanning:
   begin
   ListeActions := Criteres.StTypeActions;
   if (ListeActions <> TraduireMemoire('<<Tous>>')) and (ListeActions <> '') then
     begin
     ListeActions:=FindEtReplace(ListeActions,';','","',True);
     ListeActions:='("'+copy(ListeActions,1,Length(ListeActions)-2)+')';
     Criteres.StTypeActions := ListeActions;
     end;
   end;
end;

ChargeEtatPlanning ;
ChargeParamPlanning (TobPlannings, Planning, TobModelePlanning.Detail[0], DateSelection, DateFin, Criteres, Saisie, NumModele);
if NumModele <> TypeAgenda then BEnvoieOutlook.visible := false;
end;

Procedure TRT_Planning.MajEvenements (Var MajPlanning:THPlanning);
Begin
  MajPlanning.OnModifyItem := Modification;
  MajPlanning.OnDeleteItem := Suppression;   
  MajPlanning.OnDblClick := DoubleClick;
  MajPlanning.OnMoveItem := Deplacement;
{  MajPlanning.OnCopyItem := CopyItem;  }
  MajPlanning.OnInitItem := InitItem;
  MajPlanning.OnCreateItem := Creation;
  MajPlanning.OnOptionPopup := OnPopup;
  MajPlanning.OnAvertirApplication := AvertirApplication;
  MajPlanning.OnBeforeChange := P1BeforeChange;
  MajPlanning.OnLink := Link;
  MajPlanning.AddOptionPopup(1, '-');
  MajPlanning.AddOptionPopup(2, TraduireMemoire('Voir interlocuteur'));
  MajPlanning.AddOptionPopup(3, TraduireMemoire('Voir tiers'));
End;

Procedure TRT_Planning.P1BeforeChange(const Item: TOB;const LaRessource: String; const LaDateDeDebut, LaDateDeFin: TDateTime;var Cancel: Boolean);
begin
  Cancel:=False;
  if (DroitModif(Item,'M') = False) or (Item.GetValue('EXTERNE')='X') then
  begin
      Planning.DisplayOptionEtirer := False;
      Planning.DisplayOptionReduire := False;
      Planning.DisplayOptionDeplacement:=False;
      Planning.DisplayOptionCopie := False ;
  end
  else
  begin
      Planning.DisplayOptionCopie := False ;
      Planning.DisplayOptionEtirer := True;
      Planning.DisplayOptionReduire := True;
      Planning.DisplayOptionDeplacement:=True;
      if (CtrlHeure (Item,TobModelePlanning.Detail[0]) = False) or (TobModelePlanning.Detail[0].GetValue('CADENCEMENT') = '003') then
        begin
        Planning.DisplayOptionEtirer := False;
        Planning.DisplayOptionReduire := False;
        end;
  end
end;

procedure ChargeParamPlanning (var R: RecordPlanning ;var P:THPlanning ; T:TOB ; DateEnCours,DateFin:TDateTime;Criteres:RecordCritere;Mode,NumModele:Integer ) ;
begin
InitParamPlanning (P, T,Mode,NumModele);
//ChargeResPlanning (R);
ChargeResAff (R,Criteres);
ChargeItemsPlanning (R,T,DateEnCours,DateFin,Criteres,NumModele);
AffichagePlanning (R,P,DateEnCours,DateFin);
end;

procedure InitParamPlanning (var P:THPlanning; T:TOB; Mode,NumModele : integer );
var
  vStFormeGraph         : string;
begin
   case NumModele of
      TypePlanning :
      T.Putvalue('CADENCEMENT','003') ;  //journée
   end;

  // Changement date d'un item sur planning possible O/N
  P.MoveHorizontal := true;
  vStFormeGraph := T.Getvalue('FORMEGRAPHIQUE');

  If (vStFormeGraph = 'PGF')      Then P.FormeGraphique :=  pgFleche
  Else If (vStFormeGraph = 'PGB') Then P.FormeGraphique :=  pgBerceau
  Else If (vStFormeGraph = 'PGL') Then P.FormeGraphique :=  pgLosange
  Else If (vStFormeGraph = 'PGE') Then P.FormeGraphique :=  pgEtoile
  Else If (vStFormeGraph = 'PGA') Then P.FormeGraphique :=  pgRectangle
  Else If (vStFormeGraph = 'PGI') Then P.FormeGraphique :=  pgBisot
  Else P.FormeGraphique :=  pgRectangle;
  // MAJ couleur
  P.ColorSelection    := T.Getvalue('LSELECTION');
  P.ColorBackground   := T.Getvalue('LFOND');
  P.ColorOfSaturday   := T.Getvalue('LSAMEDI');
  P.ColorOfSunday     := T.Getvalue('LDIMANCHE');
  P.ColorJoursFeries  := T.Getvalue('LJOURSFERIES');
  P.ColSizeData    := T.Getvalue('LARGEURCOL');
  P.RowSizeData    := T.Getvalue('HAUTLIGNEDATA');

  P.ActiveLigneGroupeDate := True ;
  P.GestionJoursFeriesActive := True;
  // Encadrement
  P.FrameOn := True;
  P.MultiLine := True;
  P.ShowHint := True;

  P.Interval := piDemiHeure;

  if T.Getvalue('CADENCEMENT') = '005' then P.Interval := piHeure;
  if T.Getvalue('CADENCEMENT') = '003' then P.Interval := piJour;
  if (T.Getvalue('CADENCEMENT') = '005') or (T.Getvalue('CADENCEMENT') = '006') then
     begin
     P.ModeOutlook := True;
     P.ReductionItemOutlook := True;
     P.VoirToutesLesHeures := True;
     P.DecalageCelluleOutlook := True;
//     P.CumulInterval := pciJour;
//    P.JourneeDebut :=StrToDateTime('00:00');
//    P.JourneeFin := StrToDateTime('23:59');
     P.JourneeDebut :=StrToDateTime(T.Getvalue('HEUREDEBUT'));
     P.JourneeFin := StrToDateTime(T.Getvalue('HEUREFIN'));
     P.HourOfStart := StrToDateTime(T.Getvalue('HEUREDEBUT'));
     //Mise à jour du format de la date
//     P.DateFormat := 'hh:mm' ;
     P.DateFormat := 'ddd dd/mm/yy' ;
     P.VisionOutlook := voWeek;
     P.IntervalOutlook := pioDemiHeure;
     if T.Getvalue('CADENCEMENT') = '005' then P.IntervalOutlook := pioHeure;
     end
  else
     begin
     P.CumulInterval := pciMois;
     //Mise à jour du format de la date
     P.DateFormat := 'dd' ;
     end;
  P.SurBooking := True;
  P.ActiveSaturday := True;
  P.ActiveSunday := True;
  P.VisibleSaturday := True;
  P.VisibleSunday := True;
  P.DisplayOptionLiaison := False;
  P.DisplayOptionLier := False;
  P.DisplayOptionSuppressionLiaison := False;
//  P.DisplayOptionEtirer := True;
//  P.DisplayOptionReduire := True;
  P.DisplayOptionCopie := False;
  if mode = 1 then P.Autorisation := [patMove, patCreate, patDelete, patModify]
  else
    begin
    P.Autorisation              := [] ;
    P.DisplayOptionCreation     := False;
    P.DisplayOptionModification := False;
    P.DisplayOptionSuppression  := False;
    end;
  P.Align:=AlClient;
  P.TextAlign := taLeftJustify;
  // Les champs Item
  P.ChampLineID    := 'RAC_TYPEACTION';
  P.ChampdateDebut := 'RAC_DATEACTION';
  P.ChampDateFin   := 'DATEACTIONFIN';
  P.ChampEtat      := 'RAC_ETATACTION';
  P.ChampLibelle:='LIBELLE';
  P.ChampHint:='HINT';
  P.ResChampID:='RPA_TYPEACTION';
  case NumModele of
     TypeAgenda:
     begin
     P.TokenFieldColFixed := 'Champ1' ;
     P.TokenSizeColFixed  :='40';
     P.TokenAlignColFixed := 'C';
     end;
     TypePlanning :
     begin
     P.TokenFieldColFixed := 'RPA_LIBELLE;RPA_TYPEACTION' ;
     P.TokenSizeColFixed  :='140;50';
     P.TokenAlignColFixed := 'L;C';
     end;
  end;

  P.EtatChampCode  := 'REP_CODEETATACT' ;
  P.EtatChampLibelle := 'REP_LIBELLE' ;
  P.EtatChampBackGroundColor := 'REP_COULEURFOND' ;
  P.EtatChampFontColor := 'REP_COULEURFONTE' ;
  P.EtatChampFontName := 'REP_NOMFONTE' ;
  P.EtatChampFontSize:='REP_TAILLEFONTE';
  P.EtatChampFontStyle:='REP_STYLEFONTE';
end;

{$IFDEF CD}
function ChargeResPlanning (var R: RecordPlanning): Boolean ;
var Q     : TQuery;
    stSQL,ListeActions : String;
begin
stSQL := 'SELECT RPA_TYPEACTION,RPA_LIBELLE FROM PARACTIONS' ;
ListeActions := GetParamSoc('SO_RTPLAN_ACTION');
if (ListeActions <> TraduireMemoire('<<Tous>>')) and (ListeActions <> '') then
  begin
  ListeActions:=FindEtReplace(ListeActions,';','","',True);
  ListeActions:='("'+copy(ListeActions,1,Length(ListeActions)-2)+')';
  stSQL := stSQL + ' WHERE RPA_TYPEACTION IN '+ ListeActions;
  end;
Q := OpenSql(stSQL,True,-1,'',true);
if Not Q.Eof then
  begin
  R.TobRes.ClearDetail;
  R.TobRes.LoadDetailDB('','','',Q,False,True);
  Result := True;
  end
else
  begin
  Result := false;
  end;
Ferme(Q);
end;
{$ENDIF}

function ChargeResAff (var R: RecordPlanning ; Criteres: RecordCritere): Boolean ;
var Q     : TQuery;
    stSQL : String;
begin
R.TobResAff.ClearDetail;
stSQL := 'SELECT RPA_TYPEACTION,RPA_LIBELLE FROM PARACTIONS WHERE RPA_PRODUITPGI = "GRC" AND RPA_CHAINAGE = "---"' ;
if (Criteres.StTypeActions <> TraduireMemoire('<<Tous>>')) and (Criteres.StTypeActions <> '') then
  begin
  stSQL := stSQL + ' AND RPA_TYPEACTION IN '+ Criteres.StTypeActions;
  end;
stSQL := stSQL + ' ORDER BY RPA_TYPEACTION';
Q := OpenSql(stSQL,True,-1,'',true);
if Not Q.Eof then
  begin
  R.TobResAff.LoadDetailDB('','','',Q,False,True);
  Result := True;
  end
else
  begin
  Result := false;
  end;
Ferme(Q);
end;

function TRT_Planning.LoadTobEtat (T:TOB): integer;
var
  NewTob : Tob;
begin

  NewTOB := TOB.Create('fille_param', TobPlannings.TobEtats, -1);
  try
    NewTob.Dupliquer(T.detail[0],True,True);
  finally
    T.Free;
    result := 0;
  end;
end;

Procedure TRT_Planning.ChargeEtatPlanning  ;
var St,Etat     : String;
    vStream : TStream;
    i : integer;
    TobEtatAct : Tob;
begin
  TobPlannings.TobEtats.ClearDetail;

  VH_RT.TobParamEtatPlan.Load;

  if VH_RT.TobParamEtatPlan <> Nil then
  begin

  for i := 0 to VH_RT.TobParamEtatPlan.Detail.Count -1 do
    begin

      // chargement de la liste
      St := VH_RT.TobParamEtatPlan.Detail[i].GetValue('REP_PARAMS');

      // transfert dans une stream
      vStream := TStringStream.Create(GetRtfStringText(St));

      // recuperation dans une tob virtuelle TobsPlanning.TobEtats
      TOBLoadFromXMLStream(vStream,LoadTobEtat);
      TobPlannings.TobEtats.detail[i].AddChampSupValeur('REP_CODEETATACT', VH_RT.TobParamEtatPlan.Detail[i].GetValue('REP_CODEETATACT'),false);
      TobPlannings.TobEtats.detail[i].AddChampSupValeur('REP_LIBELLE', VH_RT.TobParamEtatPlan.Detail[i].GetValue('REP_LIBELLE'),false);
      Etat := TobPlannings.TobEtats.detail[i].getvalue('REP_STYLEFONTE');
      if Etat = 'fsBold' then TobPlannings.TobEtats.detail[i].Putvalue('REP_STYLEFONTE', 'G')
      else if Etat = 'fsItalic' then TobPlannings.TobEtats.detail[i].Putvalue('REP_STYLEFONTE', 'I');

      vStream.Free;
    end;
  // Création des états non paramétrés par l'utilisateur car suite à modif AGL
  // le planning plante si les états n'existent ds la TobEtats
  for i:=Low(CodeEtatAct) to High(CodeEtatAct) do
     begin
     if VH_RT.TobParamEtatPlan.FindFirst(['REP_CODEETATACT'],[CodeEtatAct[i]],TRUE) = Nil then
        begin
        TobEtatAct := TOB.Create('fille_param',TobPlannings.TobEtats,-1) ;
        TobEtatAct.AddChampSupValeur('REP_CODEETATACT',CodeEtatAct[i],false);
        TobEtatAct.AddChampSupValeur('REP_LIBELLE',CodeEtatAct[i],false);
        TobEtatAct.AddChampSupValeur('REP_COULEURFOND', 'ClWhite', false);
        TobEtatAct.AddChampSupValeur('REP_COULEURFONTE', 'ClBlack', false);
        TobEtatAct.AddChampSupValeur('REP_NOMFONTE', 'Times New Roman', false);
        TobEtatAct.AddChampSupValeur('REP_STYLEFONTE', 'G', false);
        TobEtatAct.AddChampSupValeur('REP_TAILLEFONTE', '10', false);
        end;
     end
  end;

end;

procedure ChargementsAppels (TOBSS: TOB;Intervenant: string;DateDebut,DateFin : TdateTime);
var MaRequete : string;
begin
  MaRequete := 'SELECT * '
    + 'FROM BTEVENPLAN '
    + 'LEFT JOIN BTETAT    ON (BTA_BTETAT=BEP_BTETAT) '
    + 'LEFT JOIN RESSOURCE ON (ARS_RESSOURCE=BEP_RESSOURCE) '
    + 'LEFT JOIN AFFAIRE   ON (AFF_AFFAIRE=BEP_AFFAIRE) '
    + 'LEFT JOIN TIERS     ON (T_TIERS=BEP_TIERS) '
    + 'LEFT JOIN ADRESSES  ON (ADR_REFCODE=AFF_AFFAIRE) ';

  MaRequete := MaRequete + 'WHERE ';
  MaRequete := MaRequete + 'BEP_RESSOURCE="'+Intervenant+'" ';
  MaRequete := MaRequete + ' AND ((BEP_DATEDEB>="' + UsDateTime(Datedebut) + '" '
    + 'AND BEP_DATEDEB<="' + UsDateTime(Datefin) + '") '
    + 'OR (BEP_DATEFIN>="' + USDateTime(DateDebut) + '" '
    + 'AND BEP_DATEFIN<="' + USDateTime(DateFin) + '") '
    + 'OR (BEP_DATEDEB<="' + UsDateTime(DateDebut) + '" '
    + 'AND BEP_DATEFIN>="' + UsDateTime(DateFin) + '")) '
    + 'AND BTA_ASSOSRES="X" ';
//  maRequete := MaRequete + 'AND AFF_ETATAFFAIRE="AFF" ';
  MaRequete := MaRequete+ ' ORDER BY BEP_DATEDEB';
  TOBSS.LoadDetailFromSQL(MaRequete,false);
end;

procedure ChargementsConges (TOBSS: TOB;Intervenant: string;DateDebut,DateFin : TdateTime);
  function ConstitueRequeteAbsenceConges (DateD,DateF : TdateTime) : string ;
  begin
    result := '( '+ // 1
              '(PCN_DATEDEBUTABS>="' + UsDateTime(DateD) + '" AND PCN_DATEDEBUTABS<="' + UsDateTime(DateF) + '") OR '+
              '(PCN_DATEFINABS>="' + USDateTime(DateD) + '" AND PCN_DATEFINABS<="' + USDateTime(DateF) + '") OR ' +
              '(PCN_DATEDEBUTABS<="' + UsDateTime(DateD) + '" AND PCN_DATEFINABS>="' + UsDateTime(DateF) + '")' +
              ')'; //1
  end;

var maRequete : string;
begin
 	MaRequete := 'SELECT *,R1.ARS_RESSOURCE FROM ABSENCESALARIE '+
  						 'LEFT JOIN RESSOURCE R1 ON R1.ARS_SALARIE=PCN_SALARIE '+
  						 'WHERE '+
  						 '(R1.ARS_RESSOURCE ="'+Intervenant+'") AND ('+
               ConstitueRequeteAbsenceConges (DateDebut,DateFin) + ') AND (PCN_DATEANNULATION="'+USDATETIME(Idate1900)+'") '+
               'AND (PCN_VALIDRESP<>"")';
  TOBSS.LoadDetailFromSQL(MaRequete,false);
end;


procedure AjouteEvtsConges (T,TOBD,TOBPlan,TOBress : TOB);
var TOBP : TOB;
		LibAffaire,Libaction,Libhint : string;
    DateD,DateF,HeureD,heureF : TdateTime;
begin
	TOBP := TOB.Create('ACTIONS',TOBPlan,-1);
  TOBP.AddChampSupValeur ('T_LIBELLE', '');
  TOBP.AddChampSupValeur ('T_VILLE', '');
  TOBP.AddChampSupValeur ('T_NATUREAUXI', '');
  TOBP.AddChampSupValeur ('HINT','');
  TOBP.AddChampSupValeur ('DATEACTIONFIN','');
  TOBP.AddChampSupValeur ('LIBELLE','');
  TOBP.AddChampSupValeur ('CODEEVENT','');
  TOBP.AddChampSupValeur ('EXTERNE','');
  LibAction := rechDom('PGMOTIFABSENCE',TOBD.getValue('PCN_TYPECONGE'),false);
  TOBP.PutValue ('CODEEVENT','');
  TOBP.putValue('RAC_TIERS','');
  TOBP.putValue('RAC_TYPEACTION','$$2');
  TOBP.putValue('RAC_ETATACTION','REA');
  TOBP.putValue('LIBELLE',LibAction);
  TOBP.putValue('HINT',LibAction);
  //
  DateD := StrToDate(DateToStr(TOBD.getValue('PCN_DATEDEBUTABS')));
  if TOBD.getValue('PCN_DEBUTDJ')='MAT' then
  begin
    heureD := GetDebutMatinee;
  end else
  begin
    heureD := GetDebutApresMidi;
  end;
  DateF := StrToDate(DateToStr(TOBD.getValue('PCN_DATEFINABS')));
  if TOBD.getValue('PCN_FINDJ')='MAT' then
  begin
    heureF := GetFinMatinee;
  end else
  begin
    heureF := GetFinApresMidi;
  end;
  //
  TOBP.putValue('RAC_DATEACTION',DateD+heureD);

  TOBP.putValue('RAC_DUREEACTION',round(CalculDureeEvenement(DateD,DateF)));
  TOBP.putValue('RAC_INTERVENANT',TOBD.getValue('ARS_RESSOURCE'));
  TOBP.putValue('DATEACTIONFIN',DateF+heureF);
	ReajustementDates (TOBP,T);
  TOBP.putValue('EXTERNE','X');
  TobP := TobRess.FindFirst(['RPA_TYPEACTION'],[TOBP.getValue('RAC_TYPEACTION')],True);
  if Tobp = Nil then
  begin
  	Tobp := Tob.create('',TobResS,-1);
    Tobp.AddChampSupValeur ('RPA_TYPEACTION',TOBP.getValue('RAC_TYPEACTION'),False);
    Tobp.AddChampSupValeur ('RPA_LIBELLE',LibAction,False);
  end;
end;

procedure AjouteEvtsAppels (T,TOBD,TOBPlan,TOBress : TOB);
var TOBP : TOB;
		LibAffaire,Libaction,Libhint : string;
begin
	TOBP := TOB.Create('ACTIONS',TOBPlan,-1);
  TOBP.AddChampSupValeur ('T_LIBELLE', TOBD.getValue('T_LIBELLE'));
  TOBP.AddChampSupValeur ('T_VILLE', TOBD.getValue('T_VILLE'));
  TOBP.AddChampSupValeur ('T_NATUREAUXI', TOBD.getValue('T_NATUREAUXI'));
  TOBP.AddChampSupValeur ('HINT','');
  TOBP.AddChampSupValeur ('DATEACTIONFIN','');
  TOBP.AddChampSupValeur ('LIBELLE','');
  TOBP.AddChampSupValeur ('CODEEVENT','');
  TOBP.AddChampSupValeur ('EXTERNE','');
  LibAction := RechDom('BTTYPEACTION',TOBD.getValue('BEP_BTETAT'),false);
  if TOBD.getValue('BEP_AFFAIRE')<>'' then
  begin
  	TOBP.PutValue ('CODEEVENT',TOBD.getValue('BEP_CODEEVENT'));
  	TOBP.putValue('RAC_TIERS',TOBD.getValue('T_TIERS'));
  	TOBP.putValue('RAC_TYPEACTION','$$$');
    if TOBD.getValue('AFF_ETATAFFAIRE')<>'AFF' then
    begin
      TOBP.putValue('RAC_ETATACTION','REA');
    end else
    begin
      TOBP.putValue('RAC_ETATACTION','PRE');
    end;
    LibAffaire := BTPCodeAffaireAffiche(TOBD.getValue('BEP_AFFAIRE'),'');
    TOBP.putValue('LIBELLE','Intervention');
    Libhint := 'Intervention N° '+LibAffaire+'#13#10'+
    					 'Client  : '+TOBD.getValue('T_LIBELLE')+ '#13#10' +
               'Adresse : '+TOBD.GetValue('ADR_ADRESSE1')+'#13#10';
    if TOBD.GetValue('ADR_ADRESSE2') <> '' then
    begin
    	LibHint := LiBHint + '          '+TOBD.GetValue('ADR_ADRESSE2')+'#13#10';
    end ;
    if TOBD.GetValue('ADR_ADRESSE3') <> '' then
    begin
    	LibHint := LibHint + '          '+TOBD.GetValue('ADR_ADRESSE3')+'#13#10';
    end;
    if TOBD.GetValue('ADR_CODEPOSTAL') <> '' then
		begin
    	LibHint := LibHint + '          '+TOBD.GetValue('ADR_CODEPOSTAL')+' '+TOBD.GetValue('ADR_VILLE')+'#13#10';
    end;
    TOBP.putValue('HINT',LibHint);
  end else
  begin
  	TOBP.PutValue('CODEEVENT',TOBD.getValue('BEP_CODEEVENT'));
  	TOBP.putValue('RAC_TYPEACTION','$$1');
    TOBP.putValue('RAC_ETATACTION','PRE');
    TOBP.putValue('LIBELLE',LibAction);
    TOBP.putValue('HINT',GetDescriptifCourt(TOBD));
  end;
  TOBP.putValue('RAC_DATEACTION',StrToDate(DateToStr(TOBD.getValue('BEP_DATEDEB')))+
  															 StrToTime(timeToStr(TOBD.getValue('BEP_HEUREDEB'))));

  TOBP.putValue('RAC_DUREEACTION',TOBD.getValue('BEP_DUREE'));
  TOBP.putValue('RAC_INTERVENANT',TOBD.getValue('BEP_RESSOURCE'));
  TOBP.putValue('DATEACTIONFIN',StrToDate(DateToStr(TOBD.getValue('BEP_DATEFIN')))+
  															 StrToTime(timeToStr(TOBD.getValue('BEP_HEUREFIN'))));
	ReajustementDates (TOBP,T);
  TOBP.putValue('EXTERNE','X');
  TobP := TobRess.FindFirst(['RPA_TYPEACTION'],[TOBP.getValue('RAC_TYPEACTION')],True);
  if Tobp = Nil then
  begin
  	Tobp := Tob.create('',TobResS,-1);
    Tobp.AddChampSupValeur ('RPA_TYPEACTION',TOBP.getValue('RAC_TYPEACTION'),False);
    Tobp.AddChampSupValeur ('RPA_LIBELLE',LibAction,False);
  end;
end;

Procedure ChargeItemsPlanning (var R: RecordPlanning; T:TOB ; DateDeb,DateFin:TDateTime; Criteres : RecordCritere; NumModele : integer ) ;
var Q         : TQuery;
    stSQL,StLibelle : String;
    i     : integer;
    Heureact : TDateTime;
    Dureeact : Double;
    recupheure:double;
    TobTemp,TOBSS : TOB;
    Heures,Minutes : integer;
begin
TOBSS := TOB.Create ('LES ELTS',nil,-1);
R.TobItems.ClearDetail;
R.TobRes.ClearDetail;

StSql:='SELECT RAC_TIERS,RAC_AUXILIAIRE,RAC_LIBELLE,RAC_NUMACTION,RAC_TYPEACTION,RAC_DATEACTION'+
               ',RAC_HEUREACTION,RAC_INTERVENANT,RAC_ETATACTION,RAC_DUREEACTION,RAC_BLOCNOTE'+
               ',RAC_CHAINAGE,RAC_NUMLIGNE,RAC_GESTRAPPEL,RAC_DELAIRAPPEL,RAC_DATERAPPEL,RAC_NUMEROCONTACT'+
               ',T_LIBELLE,T_VILLE,T_NATUREAUXI FROM RTVACTIONS WHERE ' +
               'RAC_ETATACTION in ("PRE","REA","NRE") AND ' +
               'RAC_DATEACTION >= "'+UsDateTime(DateDeb) + '" AND ' +
               'RAC_DATEACTION <= "'+UsDateTime(DateFin) + '"';
CompleteRequete (StSql,Criteres);
StSql := StSql + RTXXWhereConfident('CON');
case NumModele of
   TypeAgenda :
   StSql := StSql + 'ORDER BY RAC_DATEACTION,RAC_HEUREACTION';
   TypePlanning :
   StSql := StSql + 'ORDER BY RAC_DATEACTION';
end;
Q := OpenSql(stSQL,True,-1,'',true);
R.TobItems.LoadDetailDB('ACTIONS','','',Q,False);
ferme(Q);
if R.TobItems.detail.count > 0 then
begin
  R.TobItems.detail[0].AddChampSup('HINT',true);
  R.TobItems.detail[0].AddChampSup('DATEACTIONFIN',true);
  R.TobItems.detail[0].AddChampSup('LIBELLE',true);
  R.TobItems.detail[0].AddChampSup('EXTERNE',true);
  R.TobItems.detail[0].AddChampSupValeur ('CODEEVENT','-');

  for i:=0 to R.TobItems.detail.count-1 do
    begin
    Heureact := 0;
    if R.TobItems.detail[i].getValue('RAC_HEUREACTION') <> IDate1900 then Heureact := StrToTime(R.TobItems.detail[i].getValue('RAC_HEUREACTION'));
    R.TobItems.detail[i].PutValue('RAC_DATEACTION',R.TobItems.detail[i].GetValue('RAC_DATEACTION')+ HeureAct);

    Dureeact := R.TobItems.detail[i].getValue('RAC_DUREEACTION');
    Heures:=trunc(Dureeact/60);
    Minutes:=trunc(Dureeact-(Heures*60));
    R.TobItems.detail[i].PutValue('DATEACTIONFIN',R.TobItems.detail[i].GetValue('RAC_DATEACTION')+EncodeTime(Heures,Minutes,0,0));
//    if Dureeact <> 0 then
//       begin
       if T.GetValue('CADENCEMENT') = '005' then R.TobItems.detail[i].PutValue('DATEACTIONFIN',R.TobItems.detail[i].GetValue('DATEACTIONFIN')-0.0416666666667);
       if T.GetValue('CADENCEMENT') = '006' then R.TobItems.detail[i].PutValue('DATEACTIONFIN',R.TobItems.detail[i].GetValue('DATEACTIONFIN')-0.0208333333334);
//       end;

    StLibelle := '';
    if Criteres.StTiers = '' then StLibelle := Trim(R.TobItems.detail[i].getValue('T_LIBELLE')) + '#13#10';
    StLibelle := StLibelle + R.TobItems.detail[i].getValue('RAC_LIBELLE');
    if (CtrlHeure (R.TobItems.detail[i],T) = False) or (T.GetValue('CADENCEMENT') = '003') then
      begin
         if T.GetValue('CADENCEMENT') = '005' then recupheure := TimetoFloat (R.TobItems.detail[i].getvalue ('DATEACTIONFIN')+0.0416666666667)
           else if T.GetValue('CADENCEMENT') = '006' then recupheure := TimetoFloat (R.TobItems.detail[i].getvalue ('DATEACTIONFIN')+0.0208333333334)
             else recupheure := TimetoFloat (R.TobItems.detail[i].getvalue ('DATEACTIONFIN'));
      StLibelle := StLibelle + '#13#10' + FormatDateTime('hh:nn',R.TobItems.detail[i].getValue('RAC_HEUREACTION')) + ' - ' + FloatToStrTime(recupheure,'hh:nn');
      end;
    R.TobItems.detail[i].PutValue('LIBELLE',StLibelle);
    R.TobItems.detail[i].PutValue('HINT',R.TobItems.detail[i].getValue('RAC_LIBELLE'));
// Réajustement des dates début et fin (arrondies à l'heure ou demi-heure) en mode OUTLOOK
    if NumModele = TypeAgenda then ReajustementDates (R.TobItems.detail[i],T);
// Création de la tob des ressources en fct des items
    TobTemp := R.TobRes.FindFirst(['RPA_TYPEACTION'],[R.TobItems.detail[i].getValue('RAC_TYPEACTION')],True);
    if TobTemp = Nil then
      begin
      TobTemp := Tob.create('',R.TobRes,-1);
      TobTemp.AddChampSupValeur ('RPA_TYPEACTION',R.TobItems.detail[i].getValue('RAC_TYPEACTION'),False);
      TobTemp.AddChampSupValeur ('RPA_LIBELLE',RechDom('RTTYPEACTION',R.TobItems.detail[i].getValue('RAC_TYPEACTION'),False),False);
      end;
    end;
end;
//else PGIBox(TraduireMemoire(TexteMessage[3]),'');
// Ajout ds la tob des ressources des types d'actions non utilisées ds la période
if R.TobResAff.detail.count > 0 then
begin
  for i:=0 to R.TobResAff.detail.count-1 do
    begin
    TobTemp := R.TobRes.FindFirst(['RPA_TYPEACTION'],[R.TobResAff.detail[i].getValue('RPA_TYPEACTION')],True);
    if TobTemp = Nil then
      begin
      TobTemp := Tob.create('',R.TobRes,-1);
      TobTemp.AddChampSupValeur ('RPA_TYPEACTION',R.TobResAff.detail[i].getValue('RPA_TYPEACTION'),False);
      TobTemp.AddChampSupValeur ('RPA_LIBELLE',R.TobResAff.detail[i].getValue('RPA_LIBELLE'),False);
      end;
    end;
end;
//
if NumModele = TypeAgenda then
begin
	ChargementsAppels (TOBSS,Criteres.StIntervenant,DateDeb,DateFin);
	for I := 0 to TOBSS.detail.count -1 do
  begin
    AjouteEvtsAppels (T,TOBSS.detail[i],R.TobItems,R.tobres);
  end;
  TOBSS.clearDetail;
	ChargementsConges (TOBSS,Criteres.StIntervenant,DateDeb,DateFin);
	for I := 0 to TOBSS.detail.count -1 do
  begin
    AjouteEvtsConges (T,TOBSS.detail[i],R.TobItems,R.tobres);
  end;
end;
//
TOBSS.free;
end;

Procedure AffichagePlanning (var R: RecordPlanning;var P: THPlanning;DateEnCours,DateFin:TDateTime);
begin
  P.Activate      := False;
  P.IntervalDebut := DateEnCours;
  P.IntervalFin   := DateFin;
  P.DateOfStart   := DateEnCours;
  P.TobItems := R.TobItems;
  P.TobEtats := R.TobEtats;
  P.TobRes   := R.TobRes;
  // Activation du planning
  P.Activate := True;
end;

procedure TRT_Planning.FormDestroy(Sender: TObject);
begin
TobPlannings.TobItems.free ;
TobPlannings.TobEtats.free ;
TobPlannings.TobRes.free ;
TobPlannings.TobResAff.free ;
TobModelePlanning.free;
Planning.free;
end;

{***********A.G.L.***********************************************
Description .. : chargement des modeles de planning
Suite ........ :
Mots clefs ... :
*****************************************************************}
function TRT_Planning.LoadModeles : boolean;
var
  vListe                : TStringList;
  vStream               : TStream;
  vTOBModelePLanningXML,TobTemp : TOB; // table des modeles (blob en xml)
  ListeActions,CodeParam : string;
Begin

   result := true;
  if NumModele = 0 then CodeParam := 'PL1'
  else CodeParam := 'PL2';

  //création du modele
  vListe := TStringList.Create;

  VH_RT.TobParamPlanning.Load;

  vTOBModelePLanningXML:=VH_RT.TobParamPlanning.FindFirst(['RPP_CODEPARAMPLAN'],[CodeParam],TRUE) ;
  if vTOBModelePLanningXML <> Nil then
  begin

        // chargement de la liste
        vListe.Text := vTOBModelePLanningXML.GetValue('RPP_PARAMS');

        // transfert dans une stream
        vStream := TStringStream.Create(vListe.Text);

        // recuperation dans une tob virtuelle fTOBModelePlanning
        TOBLoadFromXMLStream(vStream, LoadTobParam);
        vStream.Free;
  end
  else
  begin
    Result := False;
    TobTemp := Tob.create('',TOBModelePlanning,-1);
    TobTemp.AddChampSupValeur('FORMEGRAPHIQUE', 'PGA', false);
    TobTemp.AddChampSupValeur('LARGEURCOL', '80', false);
    TobTemp.AddChampSupValeur('HAUTLIGNEDATA', '20', false);
    TobTemp.AddChampSupValeur('CADENCEMENT', '006', false);
    TobTemp.AddChampSupValeur('HEUREDEBUT', '08:00', false);
    TobTemp.AddChampSupValeur('HEUREFIN', '18:00', false);
    TobTemp.AddChampSupValeur('TYPEACTIONS', TraduireMemoire('<<Tous>>'), false);
    TobTemp.AddChampSupValeur('LFOND', StringToColor('clWhite'), false);
    TobTemp.AddChampSupValeur('LSELECTION', StringToColor('16777088'), false);
    TobTemp.AddChampSupValeur('LSAMEDI', StringToColor('33023'), false);
    TobTemp.AddChampSupValeur('LDIMANCHE', StringToColor('33023'), false);
    TobTemp.AddChampSupValeur('LJOURSFERIES', StringToColor('33023'), false);
  end;
  ListeActions := TOBModelePlanning.detail[0].Getvalue('TYPEACTIONS');
  if (ListeActions <> TraduireMemoire('<<Tous>>')) and (ListeActions <> '') then
     begin
     ListeActions:=FindEtReplace(ListeActions,';','","',True);
     ListeActions:='("'+copy(ListeActions,1,Length(ListeActions)-2)+')';
     TOBModelePlanning.detail[0].Putvalue('TYPEACTIONS',ListeActions)
     end;
  vListe.Free;

End;

function TRT_Planning.LoadTobParam (T:TOB): integer;
var
  NewTob : Tob;
begin

  NewTOB := TOB.Create('fille_param', TOBModelePlanning, -1);
  try
    NewTob.Dupliquer(T.detail[0],True,True);
    result := 0;
  finally
    T.Free;
  end;
end;


procedure TRT_Planning.BCalendrierClick(Sender: TObject);
var Key: Char;
    DateEnCours,DateDebut : TDateTime;
    Year,Month,Day : Word;
    NumSem  : Integer;
    vInYear : Integer;

begin
  DateEnCours := Planning.DateOfStart;
  DateEdit.text := DateToStr(DateEnCours);
  Key:='*';
  DateEdit.Enabled := True;
  PARAMDATE(Self, DateEdit, Key);
  DateEdit.Enabled := False;
  DecodeDate(StrToDate(DateEdit.text), Year, Month, Day);
  //C.B 14/10/2004
  // séparer les 2 fonctions
  NumSem := NumSemaine(StrToDate(DateEdit.text), vInYear);
  DateDebut := PremierJourSemaine(NumSem, vInYear);
  if (DateEnCours <> DateDebut) then
  begin
  ChargeItemsPlanning (TobPlannings,TobModelePlanning.detail[0],DateDebut,DateDebut+6,Criteres,NumModele);
  AffichagePlanning (TobPlannings, Planning,  DateDebut,DateDebut+6);
  end;
end;

procedure TRT_Planning.BPageSuivClick(Sender: TObject);
var DateEnCours: TDateTime;
begin
  DateEnCours := PlusDate(Planning.DateOfStart,7,'J');
  DateEdit.Text:=DateTimeToStr(DateEnCours);
  ChargeItemsPlanning (TobPlannings,TobModelePlanning.detail[0],DateEnCours,DateEnCours+6,Criteres,NumModele);
  AffichagePlanning (TobPlannings, Planning,  DateEnCours,DateEnCours+6);
end;

procedure TRT_Planning.BPagePrecClick(Sender: TObject);
var DateEnCours: TDateTime;
begin
  DateEnCours := PlusDate(Planning.DateOfStart,(-7),'J');
  DateEdit.Text:=DateTimeToStr(DateEnCours);
  ChargeItemsPlanning (TobPlannings,TobModelePlanning.detail[0],DateEnCours,DateEnCours+6,Criteres,NumModele);
  AffichagePlanning (TobPlannings, Planning,  DateEnCours,DateEnCours+6);
end;

procedure TRT_Planning.BQuitterClick(Sender: TObject);
begin
close;
end;

procedure TRT_Planning.BimprimerClick(Sender: TObject);
begin
  Planning.TypeEtat:='E';
  Planning.NatureEtat:='RPL';
  Planning.CodeEtat:='PLA';
  Planning.CriteresToPrint:='TITREPLANNING='+Caption ;
  Planning.Print(caption);
end;

procedure TRT_Planning.RESSOURCEElipsisClick(Sender: TObject);
begin
DispatchRecherche(Ressource, 3, '', '', '');
if (Ressource.text <> '') and (Ressource.text <> Criteres.StIntervenant) then
  begin
  Criteres.StIntervenant := Ressource.text;
  Caption := 'Agenda de '+ RechDom ('RTREPRESENTANT1',Criteres.StIntervenant,False);
  UpdateCaption (Self);
  ChargeItemsPlanning (TobPlannings,TobModelePlanning.detail[0],Planning.DateOfStart,Planning.DateOfStart+6,Criteres,NumModele);
  AffichagePlanning (TobPlannings, Planning,  Planning.DateOfStart,Planning.DateOfStart+6);
  end;
end;

procedure TRT_Planning.InitItem(Sender: TObject; var Item: TOB; var Cancel: boolean);
begin
     if Item = nil then Item := TOB.Create('ACTIONS', nil, -1);
     with Item do
     begin
          AddChampSup('T_LIBELLE', False);
          AddChampSup('T_VILLE', False);
          AddChampSup('T_NATUREAUXI', False);
          AddChampSup('HINT',False);
          AddChampSup('DATEACTIONFIN',False);
          AddChampSup('LIBELLE',False);
          AddChampSupValeur('CODEEVENT','');
          AddChampSupValeur('EXTERNE','-');
     end;
     Cancel := False;
end;

procedure TRT_Planning.Creation(Sender: TObject; Item: TOB; var Cancel: boolean);
var resultat,auxiliaire,select,StLibelle : string;
    NumAct : Integer;
    Q : TQUERY;
    Dureeact : Double;
    recupheure,recupdate:double;
    Heures,Minutes : word;
begin
recupheure := TimetoFloat (Item.getvalue ('RAC_DATEACTION'));
recupdate := DatetoFloat (Item.getvalue ('RAC_DATEACTION'));
if TobModelePlanning.detail[0].GetValue('CADENCEMENT') = '003' then Item.PutValue('RAC_HEUREACTION', IDate1900)
else Item.PutValue('RAC_HEUREACTION', FloatToTime(recupheure));
Item.PutValue('RAC_DATEACTION', (recupdate));
Item.PutValue ('RAC_INTERVENANT', Criteres.StIntervenant);
Item.PutValue ('RAC_TIERS', Criteres.StTiers);
TheTob := item;

Cancel := True;
Resultat :=    AGLLanceFiche('RT','RTACTIONS','','','ACTION=CREATION;CREATPLANNING') ;
if resultat <> '' then
   begin
   auxiliaire:=Trim(ReadTokenSt(resultat)) ;
   NumAct:=Valeuri(Trim(ReadTokenSt(resultat))) ;
   Select := 'SELECT RAC_TIERS,RAC_AUXILIAIRE,RAC_LIBELLE,RAC_NUMACTION,RAC_TYPEACTION,RAC_DATEACTION,RAC_HEUREACTION'+
             ',RAC_BLOCNOTE,RAC_INTERVENANT,RAC_ETATACTION,RAC_DUREEACTION,T_LIBELLE,T_VILLE,T_NATUREAUXI'+
             ',RAC_CHAINAGE,RAC_NUMLIGNE,RAC_GESTRAPPEL,RAC_DELAIRAPPEL,RAC_DATERAPPEL,RAC_NUMEROCONTACT'+
             ' FROM RTVACTIONS WHERE'+
             ' RAC_AUXILIAIRE = "'+ AUXILIAIRE+'" AND RAC_NUMACTION ='+intTostr(numact)+
             ' AND RAC_ETATACTION IN ("PRE","REA","NRE")  ' ;
   CompleteRequete (Select,Criteres);
   Select := Select + RTXXWhereConfident('CON');
   Q:=OpenSQL(Select, True,-1,'',true);
   if not Q.Eof then
      begin
      Item.PutValue ('RAC_TIERS', Q.FindField('RAC_TIERS').AsString);
      Item.PutValue ('RAC_AUXILIAIRE', auxiliaire);
      Item.PutValue ('RAC_LIBELLE', Q.FindField('RAC_LIBELLE').AsString);
      Item.PutValue ('RAC_ETATACTION', Q.FindField('RAC_ETATACTION').AsString);
      Item.PutValue ('RAC_TYPEACTION', Q.FindField('RAC_TYPEACTION').AsString);
      Item.PutValue ('RAC_DATEACTION', Q.FindField('RAC_DATEACTION').AsDateTime+Q.FindField('RAC_HEUREACTION').AsDateTime);
      Item.PutValue ('RAC_HEUREACTION', Q.FindField('RAC_HEUREACTION').AsDateTime);
      Item.PutValue ('RAC_NUMACTION', NumAct);
      Item.PutValue ('RAC_BLOCNOTE', Q.FindField('RAC_BLOCNOTE').AsString);
      Item.PutValue ('RAC_INTERVENANT', Q.FindField('RAC_INTERVENANT').AsString);
      Item.PutValue ('RAC_DUREEACTION', Q.FindField('RAC_DUREEACTION').AsFloat);
      Item.PutValue ('RAC_NUMEROCONTACT', Q.FindField('RAC_NUMEROCONTACT').AsInteger);
      Item.PutValue ('T_LIBELLE', Q.FindField('T_LIBELLE').AsString);
      Item.PutValue ('T_VILLE', Q.FindField('T_VILLE').AsString);
      Item.PutValue ('T_NATUREAUXI', Q.FindField('T_NATUREAUXI').AsString);
      Dureeact := Q.FindField('RAC_DUREEACTION').AsFloat;
      Heures:=trunc(Dureeact/60);
      Minutes:=trunc(Dureeact-(Heures*60));
      Item.PutValue ('DATEACTIONFIN',Item.GetValue('RAC_DATEACTION')+EncodeTime(Heures,Minutes,0,0));
      if TobModelePlanning.detail[0].GetValue('CADENCEMENT') = '005' then Item.PutValue('DATEACTIONFIN',Item.GetValue('DATEACTIONFIN')-0.0416666666667);
      if TobModelePlanning.detail[0].GetValue('CADENCEMENT') = '006' then Item.PutValue('DATEACTIONFIN',Item.GetValue('DATEACTIONFIN')-0.0208333333334);
      StLibelle := '';
      if Criteres.StTiers = '' then StLibelle := Trim(Item.getValue('T_LIBELLE')) + '#13#10';
      StLibelle := StLibelle + Item.getValue('RAC_LIBELLE') ;
      if (CtrlHeure (Item,TobModelePlanning.detail[0]) = False) or (TobModelePlanning.detail[0].GetValue('CADENCEMENT') = '003') then
        begin
        if TobModelePlanning.detail[0].GetValue('CADENCEMENT') = '005' then recupheure := TimetoFloat (Item.getvalue ('DATEACTIONFIN')+0.0416666666667)
          else if TobModelePlanning.detail[0].GetValue('CADENCEMENT') = '006' then recupheure := TimetoFloat (Item.getvalue ('DATEACTIONFIN')+0.0208333333334)
            else recupheure := TimetoFloat (Item.getvalue ('DATEACTIONFIN'));
        StLibelle := StLibelle + '#13#10' + FormatDateTime('hh:nn',Item.getValue('RAC_HEUREACTION')) + ' - ' + FloatToStrTime(recupheure,'hh:nn');
        end;
      Item.PutValue('LIBELLE',StLibelle);
      Item.PutValue('HINT',Item.getValue('RAC_LIBELLE'));
// Réajustement des dates début et fin (arrondies à l'heure ou demi-heure) en mode OUTLOOK
      if NumModele = TypeAgenda then ReajustementDates (Item,TobModelePlanning.detail[0]);
      Planning.InvalidateItem(Item);
      Cancel := False;
      end;
   Ferme(Q) ;
   end;
TheTob := nil;

end;

procedure TRT_Planning.Modification(Sender: TObject; Item: TOB;
  var Cancel: Boolean);
var resultat,auxiliaire,select,StLibelle : string;
    NumAct,i : Integer;
    Q : TQUERY;
    TobAction,TobItem : Tob;
    Dureeact : double;
    recupheure:double;
    Heures,Minutes : word;
    Heureact : TDateTime;
begin
Cancel := True;
if DroitModif(Item,'M') then
begin
	if Item.getValue('EXTERNE') <> 'X' then
  begin
    Auxiliaire:= Item.GetValue('RAC_AUXILIAIRE');
    NumAct:= Item.GetValue('RAC_NUMACTION');
    resultat :=   AGLLanceFiche('RT','RTACTIONS','',Auxiliaire+';'+IntToStr(NumAct),'ACTION=MODIFICATION;MODIFPLANNING') ;
    if resultat <> '' then
       begin
       if (pos('DELETE',resultat) > 0) then
         begin
         TobAction := Planning.TobItems.FindFirst(['RAC_AUXILIAIRE','RAC_NUMACTION'],[Auxiliaire,NumAct],True);
         Planning.DeleteItem (TobAction);
         end
       else
         begin
         TobAction := Planning.TobItems.FindFirst(['RAC_AUXILIAIRE','RAC_NUMACTION'],[Auxiliaire,NumAct],True);
         Planning.DeleteItem (TobAction);
         Select := 'SELECT RAC_TIERS,RAC_AUXILIAIRE,RAC_LIBELLE,RAC_NUMACTION,RAC_TYPEACTION,RAC_DATEACTION,RAC_HEUREACTION'+
                   ',RAC_BLOCNOTE,RAC_INTERVENANT,RAC_ETATACTION,RAC_DUREEACTION,T_LIBELLE,T_VILLE,T_NATUREAUXI'+
                   ',RAC_CHAINAGE,RAC_NUMLIGNE,RAC_GESTRAPPEL,RAC_DELAIRAPPEL,RAC_DATERAPPEL,RAC_NUMEROCONTACT'+
                   ' FROM ACTIONS LEFT JOIN TIERS ON RAC_TIERS = T_TIERS WHERE'+
                   ' RAC_AUXILIAIRE = "'+ AUXILIAIRE+'" AND RAC_NUMACTION ='+intTostr(numact)+
                   ' AND RAC_ETATACTION IN ("PRE","REA","NRE") ';
         CompleteRequete (Select,Criteres);
         Q:=OpenSQL(Select, True,-1,'',true);
         if not Q.Eof then
            begin
            TobItem := TOB.Create('LES_ACTIONS', nil, -1);
            TobItem.LoadDetailDB('ACTIONS','','',Q,False);
            if TobItem.detail.count > 0 then
            begin
              TobItem.detail[0].AddChampSup('HINT',true);
              TobItem.detail[0].AddChampSup('DATEACTIONFIN',true);
              TobItem.detail[0].AddChampSup('LIBELLE',true);
              for i:=0 to TobItem.detail.count-1 do
                begin
                Heureact := 0;
                if TobItem.detail[i].getValue('RAC_HEUREACTION') <> IDate1900 then
                  Heureact := StrToTime(TobItem.detail[i].getValue('RAC_HEUREACTION'));
                TobItem.detail[i].PutValue('RAC_DATEACTION',TobItem.detail[i].GetValue('RAC_DATEACTION')+ HeureAct);
                Dureeact := TobItem.detail[i].getValue('RAC_DUREEACTION');
                Heures:=trunc(Dureeact/60);
                Minutes:=trunc(Dureeact-(Heures*60));
                TobItem.detail[i].PutValue('DATEACTIONFIN',TobItem.detail[i].GetValue('RAC_DATEACTION')+EncodeTime(Heures,Minutes,0,0));
                if TobModelePlanning.detail[0].GetValue('CADENCEMENT') = '005' then
                  TobItem.detail[i].PutValue('DATEACTIONFIN',TobItem.detail[i].GetValue('DATEACTIONFIN')-0.0416666666667);
                if TobModelePlanning.detail[0].GetValue('CADENCEMENT') = '006'
                  then TobItem.detail[i].PutValue('DATEACTIONFIN',TobItem.detail[i].GetValue('DATEACTIONFIN')-0.0208333333334);
                if Criteres.StTiers = '' then StLibelle := Trim(TobItem.detail[i].getValue('T_LIBELLE')) + '#13#10' ;
                StLibelle := StLibelle + TobItem.detail[i].getValue('RAC_LIBELLE');
                if (CtrlHeure (TobItem.detail[i],TobModelePlanning.detail[0]) = False) or (TobModelePlanning.detail[0].GetValue('CADENCEMENT') = '003') then
                  begin
                  if TobModelePlanning.detail[0].GetValue('CADENCEMENT') = '005' then recupheure := TimetoFloat (TobItem.detail[i].getvalue ('DATEACTIONFIN')+0.0416666666667)
                    else if TobModelePlanning.detail[0].GetValue('CADENCEMENT') = '006' then recupheure := TimetoFloat (TobItem.detail[i].getvalue ('DATEACTIONFIN')+0.0208333333334)
                      else recupheure := TimetoFloat (TobItem.detail[i].getvalue ('DATEACTIONFIN'));
                  StLibelle := StLibelle + '#13#10' + FormatDateTime('hh:nn',TobItem.detail[i].getValue('RAC_HEUREACTION')) + ' - ' + FloatToStrTime(recupheure,'hh:nn');
                  end;
                TobItem.detail[i].PutValue('LIBELLE',StLibelle);
                TobItem.detail[i].PutValue('HINT',TobItem.detail[i].getValue('RAC_LIBELLE'));
  // Réajustement des dates début et fin (arrondies à l'heure ou demi-heure) en mode OUTLOOK
                if NumModele = TypeAgenda then ReajustementDates (TobItem.detail[i],TobModelePlanning.detail[0]);
               end;
            end;
            TobAction := TobItem.detail[0];
            TobAction.ChangeParent (Planning.TobItems,  -1);
            Planning.AddItem (TobAction) ;
            Planning.InvalidateItem (TobAction);
            TobItem.free;
            end;
         Ferme(Q) ;
         end;
       end;
	end else
  begin

  end;
end;
end;

procedure TRT_Planning.Suppression(Sender: TObject; Item: TOB; var Cancel: boolean);
var TobRess : Tob;
    Q    : TQuery;
    i : integer;
begin
cancel := True;
if Item.getValue('EXTERNE') = 'X' then exit;
if DroitModif(Item,'S') then
begin
  if (PGIAsk(TraduireMemoire(TexteMessage[7]),'') =  mrYes) then
  begin
    if GetParamSocSecur('SO_RTGESTINFOS001',False) = True then
       ExecuteSQL('DELETE FROM RTINFOS001 where RD1_CLEDATA="'+Item.GetValue('RAC_AUXILIAIRE')+';'+IntToStr(Item.GetValue('RAC_NUMACTION'))+'"') ;
    if GetParamSocSecur('SO_RTGESTIONGED',False) then
    begin
       if ExisteSQL('SELECT RTD_DOCID FROM RTDOCUMENT WHERE RTD_TIERS="'+Item.GetValue('RAC_TIERS')+'" AND RTD_NUMACTION= '+IntToStr(Item.GetValue ('RAC_NUMACTION'))) then
          ExecuteSQl('UPDATE RTDOCUMENT SET RTD_NUMACTION = 0 WHERE RTD_TIERS= "'+Item.GetValue('RAC_TIERS')
                +'" AND RTD_NUMACTION= '+ IntToStr(Item.GetValue('RAC_NUMACTION')));
    end;
    // Destruction des enregs YPLANNING si type d'action planifiable
    if RTActEstPlanifiable (Item.GetValue('RAC_TYPEACTION')) then
    begin
      TobRess := TOB.Create ('Les LIENSACTIONS',NIL,-1);
      Q:=OpenSql('SELECT RAI_GUID FROM ACTIONINTERVENANT WHERE RAI_AUXILIAIRE = "'+Item.GetValue('RAC_AUXILIAIRE')+'" AND RAI_NUMACTION ='+IntToStr(Item.GetValue('RAC_NUMACTION')),True,-1,'',true);
      try
       TobRess.loaddetailDB('','','',Q ,false);
      finally
       ferme(Q);
      end;
      for i:=0 to TobRess.detail.count-1 do
        begin
        DeleteYPL ('RAI',TobRess.Detail[i].GetString('RAI_GUID'));
        end;
      TobRess.Free;
    end;

    ExecuteSQL('DELETE FROM ACTIONINTERVENANT where RAI_AUXILIAIRE="'+Item.GetValue('RAC_AUXILIAIRE')+'" AND RAI_NUMACTION = '+IntToStr(Item.GetValue('RAC_NUMACTION'))) ;
    Item.DeleteDB(False);
    cancel := false;
  end;
end;
end;

procedure TRT_Planning.DoubleClick(Sender: Tobject);
var Item : Tob;
    Auxiliaire : string;
    NumAct : Integer;
begin
Item := THPlanning(Sender).GetCurItem ;

if Item <> Nil Then
   begin
   if Item.getValue('EXTERNE')<>'X' then
   begin
     Auxiliaire:= Item.GetValue('RAC_AUXILIAIRE');
     NumAct:= Item.GetValue('RAC_NUMACTION');
     AGLLanceFiche('RT','RTACTIONS','',Auxiliaire+';'+IntToStr(NumAct),'ACTION=CONSULTATION') ;
   end else
   begin
   end;
   end;
end;

procedure TRT_Planning.AvertirApplication(Sender: TObject; FromItem,
          ToItem: TOB; Actions: THPlanningAction);
var i : integer;
begin
  case Actions of
    paClickRight  :
    begin
      if fromItem <> nil then
        begin
          for i := 0 to NbPoPup - 1 do
            begin
            if (i = IndPoPupContact) and (FromItem.GetValue('RAC_NUMEROCONTACT') = 0) then THPlanning(Sender).EnableOptionPopup(i + 2, false)
            else THPlanning(Sender).EnableOptionPopup(i + 2, true);
            end;
        end
      else
        begin
          for i := 0 to NbPoPup - 1 do
            THPlanning(Sender).EnableOptionPopup(i + 2, false);
        end;
    end;
  end; 
end;

procedure TRT_Planning.Link (Sender:TObject; TobSource,TobDestination:TOB; Option:THPlanningOptionLink; var Cancel:boolean);
var Q : TQuery;
    iDateDebut,iDateFin : TDateTime;
    TobRessources : Tob;
    Controle : Boolean;
    SurBooking : integer;
begin
  Cancel := False;
  if TOBSource.getValue('EXTERNE')='X' then
  begin
  	Cancel := true;
    exit;
  end;
  case Option of
    polExtend:
    begin
      if Not DureeOk (TobDestination) then
         begin
         PGIBox(TraduireMemoire(TexteMessage[4]),'');
         Cancel := True;
         end
      else
         begin
         TobRessources:=Tob.create('les LIENSACTIONS',Nil,-1) ;
         Q:=OpenSql('SELECT RAI_RESSOURCE,RAI_GUID,ARS_LIBELLE FROM ACTIONINTERVENANT LEFT JOIN RESSOURCE ON ARS_RESSOURCE = RAI_RESSOURCE WHERE RAI_AUXILIAIRE = "'+TobDestination.getvalue('RAC_AUXILIAIRE')+'" AND RAI_NUMACTION ='+IntToStr(TobDestination.GetValue('RAC_NUMACTION')),True,-1,'',true);
         try
          TobRessources.loaddetailDB('','','',Q ,false,False);
         finally
          ferme(Q);
         end;
         Controle := ControleEstLibre (TobDestination,TobRessources);
         if Controle = False then SurBooking := ConfirmeSurBooking
         else SurBooking := 0;
         if SurBooking = 0 then
           begin
             if MajTobItem (TobDestination,TobSource) then
               begin
               // Mise à jour de la table YPLANNING
               if RTActEstPlanifiable (TobDestination.GetValue('RAC_TYPEACTION')) then
                 begin
                   if TobModelePlanning.Detail[0].GetValue('CADENCEMENT') = '003' then
                     begin
                     iDateDebut := TobDestination.getvalue ('RAC_DATEACTION');
                     iDateFin := TobDestination.getvalue('DATEACTIONFIN') ;
                     end
                   else
                     begin
                     iDateDebut := TobDestination.getvalue ('RAC_DATEACTION');
                     iDateFin := TobDestination.getvalue('DATEACTIONFIN');
                     if TobModelePlanning.Detail[0].GetValue('CADENCEMENT') = '005' then iDateFin := iDateFin + 0.0416666666667;
                     if TobModelePlanning.Detail[0].GetValue('CADENCEMENT') = '006' then iDateFin := iDateFin + 0.0208333333334;
                     end;
                    RTMAJYPlanning (TobRessources,iDateDebut,iDateFin,TobDestination.getvalue('RAC_LIBELLE'),TobDestination.getvalue('RAC_ETATACTION'));
                 end;
               end;
           end  // Fin SurBooking = 0
         else
           begin
           if SurBooking = -2 then PGIBox(TraduireMemoire(TexteMessage[8]),'');
           Cancel := True;
           end;
         FreeAndNil (TobRessources);
         end;
    end;

    polReduce:
    begin
       TobRessources:=Tob.create('les LIENSACTIONS',Nil,-1) ;
       Q:=OpenSql('SELECT RAI_RESSOURCE,RAI_GUID,ARS_LIBELLE FROM ACTIONINTERVENANT LEFT JOIN RESSOURCE ON ARS_RESSOURCE = RAI_RESSOURCE WHERE RAI_AUXILIAIRE = "'+TobDestination.getvalue('RAC_AUXILIAIRE')+'" AND RAI_NUMACTION ='+IntToStr(TobDestination.GetValue('RAC_NUMACTION')),True,-1,'',true);
       try
        TobRessources.loaddetailDB('','','',Q ,false,False);
       finally
        ferme(Q);
       end;
       if MajTobItem (TobDestination,TobSource) then
         begin
         // Mise à jour de la table YPLANNING
         if RTActEstPlanifiable (TobDestination.GetValue('RAC_TYPEACTION')) then
           begin
             if TobModelePlanning.Detail[0].GetValue('CADENCEMENT') = '003' then
               begin
               iDateDebut := TobDestination.getvalue ('RAC_DATEACTION');
               iDateFin := TobDestination.getvalue('DATEACTIONFIN') ;
               end
             else
               begin
               iDateDebut := TobDestination.getvalue ('RAC_DATEACTION');
               iDateFin := TobDestination.getvalue('DATEACTIONFIN');
               if TobModelePlanning.Detail[0].GetValue('CADENCEMENT') = '005' then iDateFin := iDateFin + 0.0416666666667;
               if TobModelePlanning.Detail[0].GetValue('CADENCEMENT') = '006' then iDateFin := iDateFin + 0.0208333333334;
               end;
              RTMAJYPlanning (TobRessources,iDateDebut,iDateFin,TobDestination.getvalue('RAC_LIBELLE'),TobDestination.getvalue('RAC_ETATACTION'));
           end;
         end;
       FreeAndNil (TobRessources);
    end;
  end;
end;

procedure TRT_Planning.Deplacement (Sender: TObject; Item: TOB; var Cancel: boolean);
var Q : TQuery;
    TypeAction : string;
    iDateDebut,iDateFin : TDateTime;
    TobRessources : Tob;
    Controle : Boolean;
    SurBooking : integer;
begin
Cancel := True;
if Item.getValue('EXTERNE')='X' then exit;
Q := OpenSQL('Select RAC_TYPEACTION FROM ACTIONS WHERE RAC_AUXILIAIRE="' + Item.getvalue('RAC_AUXILIAIRE') + '" AND RAC_NUMACTION ='+intTostr(Item.GetValue('RAC_NUMACTION')), True,-1,'',true);
if not Q.Eof then TypeAction := Q.findfield('RAC_TYPEACTION').Asstring;
Ferme(Q);
if Item.Getvalue ('RAC_TYPEACTION') <> TypeAction then
   begin
   PGIBox(TraduireMemoire(TexteMessage[5]),'');
   end
else
   begin
   TobRessources:=Tob.create('les LIENSACTIONS',Nil,-1) ;
   Q:=OpenSql('SELECT RAI_RESSOURCE,RAI_GUID,ARS_LIBELLE FROM ACTIONINTERVENANT LEFT JOIN RESSOURCE ON ARS_RESSOURCE = RAI_RESSOURCE WHERE RAI_AUXILIAIRE = "'+Item.getvalue('RAC_AUXILIAIRE')+'" AND RAI_NUMACTION ='+IntToStr(Item.GetValue('RAC_NUMACTION')),True,-1,'',true);
   try
    TobRessources.loaddetailDB('','','',Q ,false,False);
   finally
    ferme(Q);
   end;
   Controle := ControleEstLibre (Item,TobRessources);
   if Controle = False then SurBooking := ConfirmeSurBooking
   else SurBooking := 0;
   if SurBooking = 0 then
     begin
       if MajTobItem (Item,Nil) then
         begin
         Cancel := False;
         // Mise à jour de la table YPLANNING
         if RTActEstPlanifiable (Item.GetValue('RAC_TYPEACTION')) then
{         VH_RT.TobTypesAction.Load;
         TobTypAct:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[Item.GetValue('RAC_TYPEACTION'),'---',0],TRUE) ;
         if (TobTypAct <> Nil) and (TobTypAct.GetBoolean('RPA_PLANIFIABLE') = True) then   }
           begin
           if TobModelePlanning.Detail[0].GetValue('CADENCEMENT') = '003' then
             begin
             iDateDebut := Item.getvalue ('RAC_DATEACTION');
             iDateFin := Item.getvalue('DATEACTIONFIN') ;
             end
           else
             begin
             iDateDebut := Item.getvalue ('RAC_DATEACTION');
             iDateFin := Item.getvalue('DATEACTIONFIN');
             if TobModelePlanning.Detail[0].GetValue('CADENCEMENT') = '005' then iDateFin := iDateFin + 0.0416666666667;
             if TobModelePlanning.Detail[0].GetValue('CADENCEMENT') = '006' then iDateFin := iDateFin + 0.0208333333334;
             end;
            RTMAJYPlanning (TobRessources,iDateDebut,iDateFin,Item.getvalue('RAC_LIBELLE'),Item.getvalue('RAC_ETATACTION'));
           end;
        end;
     end  // Fin SurBooking = 0
   else
     begin
     if SurBooking = -2 then PGIBox(TraduireMemoire(TexteMessage[8]),'');
     end;
   FreeAndNil (TobRessources);
   end;
end;

function TRT_Planning.MajTobItem(Tobitem:TOB;FromItem: TOB): Boolean;
Var DateF,DateD,Heureact :TdateTime;
    recupheure,recupdate,dureeact:double;
    StUpdate,StLibelle : string;
    Hour,Min,Sec,Msec : Word;
begin
   result := True;
   with TobItem do
   begin
     if TobModelePlanning.Detail[0].GetValue('CADENCEMENT') = '003' then
       begin
         Heureact := 0;
         if getValue('RAC_HEUREACTION') <> IDate1900 then Heureact := StrToTime(getValue('RAC_HEUREACTION'));
         PutValue('RAC_DATEACTION',GetValue('RAC_DATEACTION')+ HeureAct);
         Dureeact := getValue('RAC_DUREEACTION');
         Hour:=trunc(Dureeact/60);
         Min:=trunc(Dureeact-(Hour*60));
         PutValue('DATEACTIONFIN',GetValue('RAC_DATEACTION')+EncodeTime(Hour,Min,0,0));
         CalculDateRappel (Tobitem,FromItem);
         recupdate := DatetoFloat (getvalue ('RAC_DATEACTION'));
         StUpdate := 'UPDATE ACTIONS SET RAC_DATEACTION = "' + UsDateTime(recupdate)+'"';
         StUpdate := StUpdate + ', RAC_DATERAPPEL = "' + UsTime(GetValue('RAC_DATERAPPEL'))+'"';
         StUpdate := StUpdate + ', RAC_UTILISATEUR = "' + V_PGI.User+'"';
         StUpdate := StUpdate + ', RAC_DATEMODIF = "' + UsTime(NowH)+'"';
         StUpdate := StUpdate + ' WHERE RAC_AUXILIAIRE = "' + GetValue('RAC_AUXILIAIRE');
         StUpdate := StUpdate + '" AND RAC_NUMACTION = ' + intTostr(GetValue('RAC_NUMACTION'));
       end
     else
       begin
       DateF := getvalue ('DATEACTIONFIN');
       if TobModelePlanning.Detail[0].GetValue('CADENCEMENT') = '005' then DateF := DateF + 0.0416666666667;
       if TobModelePlanning.Detail[0].GetValue('CADENCEMENT') = '006' then DateF := DateF + 0.0208333333334;
       DateD := getvalue ('RAC_DATEACTION');
       DateF := DateF-DateD;
       DecodeTime(DateF, Hour, Min, Sec, MSec);
       Dureeact := (Hour*60)+Min;
  {     if (Dureeact > 720) then
          begin
            PGIBox('La durée ne peut pas dépasser 12 heures','');
            result := False;
          end
       else
          begin }
         PutValue('RAC_DUREEACTION', Dureeact);
         recupheure := TimetoFloat (getvalue ('RAC_DATEACTION'));
         recupdate := DatetoFloat (getvalue ('RAC_DATEACTION'));
         PutValue('RAC_HEUREACTION', FloatToTime(recupheure));
         CalculDateRappel (Tobitem,FromItem);
         StUpdate := 'UPDATE ACTIONS SET RAC_DATEACTION = "' + UsDateTime(recupdate)+'"';
         StUpdate := StUpdate + ', RAC_HEUREACTION = "' + UsTime(GetValue('RAC_HEUREACTION'))+'"';
         StUpdate := StUpdate + ', RAC_DUREEACTION = ' + FloatToStr(GetValue('RAC_DUREEACTION'));
         StUpdate := StUpdate + ', RAC_DATERAPPEL = "' + UsTime(GetValue('RAC_DATERAPPEL'))+'"';
         StUpdate := StUpdate + ', RAC_UTILISATEUR = "' + V_PGI.User+'"';
         StUpdate := StUpdate + ', RAC_DATEMODIF = "' + UsTime(NowH)+'"';
         StUpdate := StUpdate + ' WHERE RAC_AUXILIAIRE = "' + GetValue('RAC_AUXILIAIRE');
         StUpdate := StUpdate + '" AND RAC_NUMACTION = ' + intTostr(GetValue('RAC_NUMACTION'));
         end;
       if (ExecuteSql(StUpdate) <> 1) then
          begin
            PGIBox(TraduireMemoire(TexteMessage[6]),'');
            result := False;
          end
       else
          begin
          if Criteres.StTiers = '' then StLibelle := Trim(getValue('T_LIBELLE'))+ '#13#10' ;
          StLibelle := StLibelle + GetValue('RAC_LIBELLE');
          if (CtrlHeure (TobItem,TobModelePlanning.detail[0]) = False) or (TobModelePlanning.Detail[0].GetValue('CADENCEMENT') = '003') then
            begin
            if TobModelePlanning.detail[0].GetValue('CADENCEMENT') = '005' then recupheure := TimetoFloat (getvalue ('DATEACTIONFIN')+0.0416666666667)
              else if TobModelePlanning.detail[0].GetValue('CADENCEMENT') = '006' then recupheure := TimetoFloat (getvalue ('DATEACTIONFIN')+0.0208333333334)
                else recupheure := TimetoFloat (getvalue ('DATEACTIONFIN'));
            StLibelle := StLibelle + '#13#10' + FormatDateTime('hh:nn',getValue('RAC_HEUREACTION')) + ' - ' + FloatToStrTime(recupheure,'hh:nn');
            end;
          PutValue('LIBELLE',StLibelle);
          Planning.InvalidateItem(TobItem);
          end;
   {     end;  }
   end;
end;

procedure TRT_Planning.OnPopup(Item: tob; ZCode: integer; var Redraw: boolean);
begin
if ZCode = 2 then AGLLanceFiche('YY','YYCONTACT','T;'+Item.GetValue('RAC_AUXILIAIRE')+';'+IntToStr(Item.GetValue('RAC_NUMEROCONTACT')),'','ACTION=CONSULTATION') ;
if ZCode = 3 then AGLLanceFiche('GC','GCTIERS','',Item.GetValue('RAC_AUXILIAIRE'),'MONOFICHE;ACTION=CONSULTATION;T_NATUREAUXI='+Item.GetValue('T_NATUREAUXI')) ;
end;

function TRT_Planning.DroitModif(Tobitem: TOB;TypeModif: string): Boolean;
var TobTypActEncours : Tob;
    Q : TQUERY;
    Select : String;
    Err : integer;
begin
Err := 0;
Result := True;
if not RTDroitModifActions(Tobitem.GetValue('RAC_TIERS'),Tobitem.GetValue('RAC_TYPEACTION'),Tobitem.GetValue('RAC_INTERVENANT')) then
  begin
  Result := False;
  Err := 1;
  end;

VH_RT.TobTypesAction.Load;

TobTypActEncours:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[Tobitem.GetValue('RAC_TYPEACTION'),Tobitem.GetValue('RAC_CHAINAGE'),Tobitem.GetValue('RAC_NUMLIGNE')],TRUE) ;
if TobTypActEncours = Nil then
begin
    TobTypActEncours:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[Tobitem.GetValue('RAC_TYPEACTION'),'---',0],TRUE) ;
    if TobTypActEncours = Nil then exit;
end;
if (TobTypActEncours.GetValue('RPA_MODIFRESP')='X')  then
  begin
  if not RTDroitModifTypeAction (Tobitem.GetValue ('RAC_PRODUITPGI'))then
      begin
      Select := 'SELECT ARS_UTILASSOCIE FROM RESSOURCE WHERE ARS_RESSOURCE = "'+ Tobitem.GetValue ('RAC_INTERVENANT')+'"';
      Q:=OpenSQL(Select, True,-1,'RESSOURCE',true);
      if not Q.Eof then
         if Q.FindField('ARS_UTILASSOCIE').asString <> V_PGI.User then
         begin
             Result := False;
             Err := 1;
         end;
      Ferme(Q) ;
      end;
  end;
if Result = True then
  begin
  if TypeModif = 'S' then
     begin
     if TobTypActEncours.GetValue('RPA_NONSUPP') = 'X' then
        begin
        Result := False;
        Err := 2;
        end;
     end;
  end;
if Result = False then PGIBox(TraduireMemoire(TexteMessage[Err]),'');
end;

function TRT_Planning.DureeOK(Tobitem: TOB): Boolean;
Var DateF,DateD:TdateTime;
    dureeact:double;
    Hour,Min,Sec,Msec : Word;
begin
   result := True;
   with TobItem do
   begin
     DateF := getvalue ('DATEACTIONFIN');
     if TobModelePlanning.Detail[0].GetValue('CADENCEMENT') = '005' then DateF := DateF + 0.0416666666667;
     if TobModelePlanning.Detail[0].GetValue('CADENCEMENT') = '006' then DateF := DateF + 0.0208333333334;
     DateD := getvalue ('RAC_DATEACTION');
     DateF := DateF-DateD;
     if (DateF >= 1) then Dureeact := 1000
     else
         begin
         DecodeTime(DateF, Hour, Min, Sec, MSec);
         Dureeact := (Hour*60)+Min;
         end;
     if (Dureeact > 720) then result := False;
   end;
end;

function TRT_Planning.ControleEstLibre(Item,TobRessources:TOB): Boolean;
var
    DureeAct : double;
    Heures,Minutes : integer;
    i : integer;
    iDateDebut,iDateFin : TDateTime;
    TobControleYplanning,TF : Tob;
begin
Result := True;
if RTActEstPlanifiable (Item.GetValue('RAC_TYPEACTION')) then
  begin
  TobControleYplanning := Tob.create('CONTROLEYPLANNING',Nil,-1) ;
  for i:=0 to TobRessources.detail.count-1 do
    begin
    TF:=Tob.create('#CONTROLE',TobControleYplanning,-1);
    TF.AddChampSupValeur('RESSOURCE',TobRessources.Detail[i].GetString('RAI_RESSOURCE') );
    TF.AddChampSupValeur('NOMRESSOURCE',TobRessources.Detail[i].GetString('ARS_LIBELLE') );
    TF.AddChampSupValeur('GUID',TobRessources.Detail[i].GetString('RAI_GUID'));
    TF.AddChampSupValeur('LIBRE','X');
    TF.AddChampSupValeur('MOTIF','');
    end;
  if TobControleYplanning.Detail.Count <> 0 then
    begin
    if TobModelePlanning.Detail[0].GetValue('CADENCEMENT') = '003' then
      begin
      iDateDebut := Item.getvalue ('RAC_DATEACTION') + Item.getvalue ('RAC_HEUREACTION');
      Dureeact := Item.getvalue('RAC_DUREEACTION');
      Heures:=trunc(Dureeact/60);
      Minutes:=trunc(Dureeact-(Heures*60));
      iDateFin := Item.getvalue('RAC_DATEACTION') + Item.getvalue('RAC_HEUREACTION') + EncodeTime(Heures,Minutes,0,0);
      end
    else
      begin
      iDateDebut := Item.getvalue ('RAC_DATEACTION');
      iDateFin := Item.getvalue('DATEACTIONFIN');
      if TobModelePlanning.Detail[0].GetValue('CADENCEMENT') = '005' then iDateFin := iDateFin + 0.0416666666667;
      if TobModelePlanning.Detail[0].GetValue('CADENCEMENT') = '006' then iDateFin := iDateFin + 0.0208333333334;
      end;
    Result := ControleIsFreeYPL (TobControleYplanning,iDateDebut,iDateFin,True);
    end;
  FreeAndNil (TobControleYplanning);
  end;
end;
{***********A.G.L.***********************************************
Description .. : fonction qui retourne la partie entiere de la date, cad sans
Suite ........ : l'heure
*****************************************************************}

function DateToFloat(TDate: TDateTime): Double;
var tmp: double;
begin
  tmp := Tdate - 0.49999;
  result := arrondi(Tmp, 0);
end;

procedure CalculDateRappel(ToItem:TOB;FromItem: TOB);
var CalculDate: Boolean;
    DateAction : TDateTime;  // contient date+heure
begin
CalculDate := True;
if FromItem <> Nil then
  begin
  if ToItem.GetValue('RAC_DATEACTION') = FromItem.GetValue('RAC_DATEACTION') then CalculDate := False;
  end;
if ToItem.GetValue('RAC_GESTRAPPEL') <> 'X' then CalculDate := False;
if CalculDate = True then
  begin
  DateAction := ToItem.GetValue('RAC_DATEACTION');
  if ToItem.GetValue ('RAC_DELAIRAPPEL') = '' then
     ToItem.PutValue('RAC_DATERAPPEL',DateAction)
  else
    begin
      if ToItem.GetValue ('RAC_DELAIRAPPEL') < '001' then
         ToItem.PutValue('RAC_DATERAPPEL',DateAction-EncodeTime(0, StrToInt(copy(ToItem.GetValue ('RAC_DELAIRAPPEL'),2,2)), 0, 0))
      else
        if ToItem.GetValue ('RAC_DELAIRAPPEL') < '024' then
          ToItem.PutValue('RAC_DATERAPPEL',DateAction-EncodeTime(ToItem.GetValue ('RAC_DELAIRAPPEL'), 0, 0, 0))
        else
          ToItem.PutValue('RAC_DATERAPPEL',PlusDate(DateAction, (ToItem.GetValue ('RAC_DELAIRAPPEL')/24)* (-1),'J'));
    end;
  end;
end;

procedure CompleteRequete(var Requete: string; Criteres: RecordCritere);
var i : integer;
begin
if (Criteres.StIntervenant <> '') then
   begin
   i := Pos ('RTVACTIONS',Requete);
   if i <> 0 then
     begin
     Requete := Copy(Requete,1,i+10)+ 'LEFT JOIN ACTIONINTERVENANT ON RAC_AUXILIAIRE = RAI_AUXILIAIRE AND RAC_NUMACTION= RAI_NUMACTION '
              + Copy(Requete,i+10+1,strlen(Pchar(Requete)));
     Requete := Requete + ' AND RAI_RESSOURCE = "'+Criteres.StIntervenant+'"';
     end;
   end;
if (Criteres.StTiers <> '') then
   begin
   Requete := Requete + ' AND RAC_TIERS = "'+Criteres.StTiers+'" ';
   end;
if (Criteres.StTypeActions <> TraduireMemoire('<<Tous>>')) and (Criteres.StTypeActions <> '') then
   begin
   Requete := Requete + ' AND RAC_TYPEACTION IN ' + Criteres.StTypeActions;
   end;
end;

procedure ReajustementDates(Tobitem: TOB;T:TOB);
var Hour, Min, Sec, MSec: Word;
    m : Double;
begin
if T.GetValue('CADENCEMENT') = '005' then m:=0.0416666666667
else m:=0.0208333333334;
//heureact := Tobitem.GetValue('RAC_DATEACTION');
DecodeTime(Tobitem.getvalue ('RAC_DATEACTION'),Hour,Min,Sec,MSec);
if T.GetValue('CADENCEMENT') = '005' then Min := 0
  else
begin
  if Min < 30 then Min := 0
  else  Min := 30;
end;
Tobitem.PutValue('RAC_DATEACTION',Trunc (Tobitem.GetValue('RAC_DATEACTION'))+(EncodeTime(Hour,Min,0,0)));
//heureact := Tobitem.GetValue('RAC_DATEACTION');
//heureact := Tobitem.GetValue('DATEACTIONFIN');
DecodeTime(Tobitem.getvalue ('DATEACTIONFIN')+m,Hour,Min,Sec,MSec);
if T.GetValue('CADENCEMENT') = '005' then Min := 0
  else
begin
  if Min < 30 then Min := 0
  else  Min := 30;
end;
Tobitem.PutValue('DATEACTIONFIN',Trunc (Tobitem.GetValue('DATEACTIONFIN'))+(EncodeTime(Hour,Min,0,0)));
Tobitem.PutValue('DATEACTIONFIN',Tobitem.GetValue('DATEACTIONFIN')-m);
//heureact := Tobitem.GetValue('DATEACTIONFIN');
end;

procedure AGLRTExecPlanning( parms: array of variant; nb: integer );
var  F      : TForm;
     DateDeb,DateFin : TDateTime;
     Mode : Integer;
     Year,Month,Day : Word;
     NumSem  : Integer;
     vInYear : Integer;
  
begin
  F := TForm(Longint(Parms[0])) ;
  if not assigned(F) then exit;
  DateDeb := StrToDate(Parms[2]);
  if Parms[3] = '' then
     begin
     DecodeDate(DateDeb, Year, Month, Day);
     NumSem := NumSemaine(DateDeb, vInYear);
     DateDeb := PremierJourSemaine(NumSem, vInYear);
     DateFin := DateDeb+6;
     end
  else DateFin := StrToDate(Parms[3]);
  Mode := Valeuri(Parms[4]);
  ExecutePlanning (String(Parms[1]),DateDeb,DateFin,Mode,Parms[5],Parms[6]);
end;

procedure TRT_Planning.BRechargerClick(Sender: TObject);
var DateDeb,DateFin: TDateTime;
begin
  DateDeb := Planning.IntervalDebut;
  DateFin := Planning.IntervalFin;
  ChargeItemsPlanning (TobPlannings,TobModelePlanning.detail[0],DateDeb,DateFin,Criteres,NumModele);
  AffichagePlanning (TobPlannings, Planning,  DateDeb,DateFin);
end;

procedure TRT_Planning.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;

procedure TRT_Planning.RESSOURCEExit(Sender: TObject);
begin
if (Ressource.text <> '') and (Ressource.text <> Criteres.StIntervenant) then
  begin
  Criteres.StIntervenant := Ressource.text;
  Caption := 'Agenda de '+ RechDom ('RTREPRESENTANT1',Criteres.StIntervenant,False);
  UpdateCaption (Self);
  ChargeItemsPlanning (TobPlannings,TobModelePlanning.detail[0],Planning.DateOfStart,Planning.DateOfStart+6,Criteres,NumModele);
  AffichagePlanning (TobPlannings, Planning,  Planning.DateOfStart,Planning.DateOfStart+6);
  end;
end;

function GetDescriptifCourt ( TOBD : TOB) : string;
var
  Memo: Tmemo;
  Panel: TPanel;
begin
  Panel := TPanel.Create(nil);
  try
    Panel.Visible := False;
    Panel.ParentWindow := GetDesktopWindow;
    Memo := Tmemo.Create(Panel);
    Memo.Parent := Panel;
    Memo.Visible := False;
    memo.text :=  TOBD.GetValue('BEP_BLOCNOTE');
    result := memo.lines.Strings [0]+'#13#10';
    if (memo.lines.Strings [1] <> '') and (memo.lines.Strings [1]<>'#13#10') then
    begin
    	result := result + memo.lines.Strings [1];
    end;
    if (memo.lines.Strings [2] <> '') and (memo.lines.Strings [2]<>'#13#10') then
    begin
    	result := result +'#13#10...';
    end;
    Memo.Free;
  finally
    Panel.Free;
  end

end;

procedure TRT_Planning.BEnvoieOutlookClick(Sender: TObject);
var Indice : integer;
		TOBP : TOB;
begin
	for Indice := 0 to TobPlannings.TobItems.detail.count - 1 do
  begin
  (*
    if TOBP.getValue('EXTERNE')='-' then
    begin
    end else if TOBP.getValue('
  	AddTache('Tiers : '+TOBP.GetValue('RAC_TIERS')+', Action : '+getfield ('RAC_LIBELLE'),Memo.Text,getfield ('RAC_DATEACTION'),getfield ('RAC_DATEECHEANCE')) ;
  *)
  end;
end;

Initialization
  RegisterAglProc('RTExecPlanning', TRUE, 6, AGLRTExecPlanning);
end.

