{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 06/09/2004
Modifié le ... :   /  /
Description .. : TOM Gestion des sessions des stages
Mots clefs ... : FORMATION ; PLANNING
*****************************************************************
---- JL 17/10/2006 Modification contrôle des exercices de formations -----
PT1 28/09/2007  V_7  FL  Emanager / Report / Adaptation cursus + accès assistant
PT2 18/10/2007  V_7  FL  Emanager / Report / Ajout de la légende
PT3 27/10/2007  V_8  FL  Emanager / Report / Intégration du planning dans la fiche 
PT4 04/12/2007  V_8  FL  Emanager / Report / Pas de code stage, code couleur différent + correctifs
}
unit PGPlanningSessionForm;

interface

uses                        
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
{$IFDEF EAGLCLIENT}
  UtileAGL,eMul,MaineAgl,HStatus,
{$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,
{$ENDIF}
  HCtrls,StdCtrls, HTB97, ExtCtrls,hplanning,UTOB, Dialogs,
  HEnt1,HeureUtil,PGOutilsFormation,HmsgBox,HPanel,uiutil;
const
  SBWM_APPSTARTUP  = WM_USER + 1;

Type
  TNiveauRupture = record        // ENREGISTREMENT DE TYPE NIVEAU DE RUPTURE
     NiveauRupt :   0..4;   // Libellé niveau de rupture
     ChampsRupt : Array [1..4] of String;   // Nom de champs niveau de rupture
     CondRupt   : String;   // Clause conditionnel niveau de rupture
  end;


  TPlanningFormation = class(TForm)
    PPanel: TPanel;
    Dock971: TDock97;
    PBouton: TToolWindow97;
    BImprimer: TToolbarButton97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    TRECAP: TToolbarButton97;
    Bperso: TToolbarButton97;
    TExcel: TToolbarButton97;
    BExporter: TToolbarButton97;
    BtnDetail: TToolbarButton97;
    PRECAP: TPanel;
    GBRecap: TGroupBox;
    LBSTAGE: THLabel;
    HSTAGE: THLabel;
    HSESSION: THLabel;
    LBSESSION: THLabel;
    HNATURE: THLabel;
    LBNATURE: THLabel;
    HLIEU: THLabel;
    LBLIEU: THLabel;
    LBDATADEB: THLabel;
    HDATEDEB: THLabel;
    LBDATEFIN: THLabel;
    HDATEFIN: THLabel;
    GBLEGENDE: TGroupBox;
    CLR1: TLabel;
    LBLCLR1: TLabel;
    LBLCLR2: TLabel;
    LBLCLR3: TLabel;
    LBLCLR4: TLabel;
    CLR2: TLabel;
    CLR3: TLabel;
    CLR4: TLabel;
    LBLNBINSC: THLabel;
    HNBINSC: THLabel;
    LBLNBDISPO: THLabel;
    HNBDISPO: THLabel;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BFermeClick(Sender: TObject);
    procedure BtnDetailClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure TExcelClick(Sender: TObject);
    //Selection Item
    procedure HAvertirApplication(Sender: TObject; FromItem, ToItem: TOB; Actions: THPlanningAction);
    //DoubleClick sur ressource
    procedure DoubleCLickSpecPlanning(ACol,ARow: INTEGER ; TypeCellule: TPlanningTypeCellule ; T: TOB=Nil) ;
    //DoubleClick sur item
    procedure DoubleCLickPlanning(Sender: TObject);
private
    { Déclarations privées }
    H: THPlanning ;
    StWhTobEtats,StWhTobItems,StWhTobRessource,MillesimeEC : string;
    BackGroundColor,FontColor,FontName,FontStyle,ColorBackground : String;
    FontSize : Integer;
    TobRessources,TobItems,TobEtats : TOB ;
    InitDateDebFor,InitDateFinFor,DateDebFor,DateFinFor : TDateTime ;
    Rupture : TNiveauRupture;
    MultiNiveau,OnShow : Boolean;
    ChampsOrg : Array [1..4] of String;
    Champslibre : Array [1..4] of String;
    DDMillesimeEC,DFMillesimeEC : TDateTime;
    Function  ChargeTobPlaning : boolean;
//    procedure MiseEnFormeTobRupt(TypeTob : string ; Tob_Stages : Tob);
//    Function  AddRessTob(Tob_Stages : Tob) : Boolean;
//    Function  CreateEnr(Champ: string; Tob_Stages: Tob) : Boolean;
    procedure WMAppStartup(var msg: TMessage); message SBWM_APPSTARTUP;
    procedure LoadConfigPlanning;
//    function  ListeChampRupt(Sep : string) : string;
    procedure RefrehRecapPlanning(Stage,Millesime : string;Numordre:Integer; TItem : Tob);
  public
    { Déclarations publiques }
  end;
  procedure PGPlanningFormation(TP : THPanel; DateDebFor,DateFinFor : TDateTime ;StWhTobEtats,StWhTobItems,StWhTobRessource : string; NiveauRupt : TNiveauRupture; MultiNiveau : boolean = False ); //PT3
  //PT3 - Début
  procedure PGPlanningFormationDestroy;
  procedure PGPlanningFormationExport;
  procedure PGPlanningFormationShowDetail;
  procedure PGPlanningFormationImprimer;
  //PT3 - Fin

var
  PlanningFormation: TPlanningFormation;
implementation

uses EntPaie;

{$R *.DFM}

//PT3 - Début
procedure PGPlanningFormationDestroy ();
Begin
    if PlanningFormation<>nil then
    Begin
         PlanningFormation.Close;
         FreeAndNil(PlanningFormation);
    End;
End;

procedure PGPlanningFormationShowDetail;
Begin
    if PlanningFormation<>nil then PlanningFormation.BtnDetail.Click;
End;

procedure PGPlanningFormationExport;
Begin
    if PlanningFormation<>nil then PlanningFormation.TExcel.Click;
End;

procedure PGPlanningFormationImprimer;
Begin
    if PlanningFormation<>nil then PlanningFormation.BImprimer.Click;
End;
//PT3 - Fin

procedure PGPlanningFormation(TP : THPanel; DateDebFor,DateFinFor : TDateTime ;StWhTobEtats,StWhTobItems,StWhTobRessource : string; NiveauRupt : TNiveauRupture; MultiNiveau : boolean = False ); //PT3
BEGIN
  //PlanningFormation:=nil; //PT3
  If PlanningFormation <> Nil Then PGPlanningFormationDestroy;  //PT3
  PlanningFormation := TPlanningFormation.Create(Application);
  try
    PlanningFormation.DateDebFor:=DateDebFor;
    PlanningFormation.DateFinFor:=DateFinFor;
    PlanningFormation.StWhTobEtats:=StWhTobEtats;
    PlanningFormation.StWhTobItems:=StWhTobItems;
    PlanningFormation.StWhTobRessource:=StWhTobRessource;
    PlanningFormation.Rupture:=NiveauRupt;
    PlanningFormation.MultiNiveau:=MultiNiveau ;
//    {$IFNDEF EMANAGER}
//    PlanningFormation.ShowModal ;
//    {$ELSE}
    //PT3 - Début
    PlanningFormation.BorderStyle := bsNone;
    PlanningFormation.Parent :=  TP;
    PlanningFormation.Show;
    If PlanningFormation.H.Visible = False Then FreeAndNil(PlanningFormation);
    //PT3 - Fin
//    {$ENDIF}

  finally
//    {$IFNDEF EMANAGER}
//    if PlanningFormation<>nil then PlanningFormation.Free;
//    {$ENDIF}
  end;
End;

procedure TPlanningFormation.WMAppStartup(var msg: TMessage);
begin
  CLose ;
end;

procedure TPlanningFormation.FormShow(Sender: TObject);
{$IFDEF EMANAGER}var Q : TQuery;{$ENDIF}
begin
OnShow:=True;
//PT1 - Début
{$IFDEF EMANAGER}
Q := OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN,PFE_MILLESIME FROM EXERFORMATION WHERE PFE_ENCOURS="X" ORDER BY PFE_MILLESIME DESC',True);
If Not Q.Eof then
begin
        DDMillesimeEC := Q.FindField('PFE_DATEDEBUT').AsDateTime;
        DFMillesimeEC := Q.FindField('PFE_DATEFIN').AsDateTime;
        MillesimeEC := Q.FindField('PFE_MILLESIME').AsString;
end
else
begin
        DDMillesimeEC := IDate1900;
        DFMillesimeEC := IDate1900;
        MillesimeEC := '0000';
end;
Ferme(Q);
{$ELSE}
MillesimeEC := RendMillesimeRealise(DDMillesimeEC,DFMillesimeEC);
{$ENDIF}
//PT1 - Fin
HSTAGE.Caption := '';
HSESSION.Caption := '';
HNATURE.Caption := '';
HLIEU.Caption := '';
HDATEDEB.Caption := '';
HDATEFIN.Caption := '';
HNBDISPO.Caption := ''; //pt2
HNBINSC.Caption := '';  //pt2
FontColor:='ClBlack';
FontStyle :='G';
FontName:='Times New Roman';
FontSize:=10;
ColorBackground:='clWhite';
InitDateDebFor:=PlanningFormation.DateDebFor;
InitDateFinFor:=PlanningFormation.DateFinFor;
ChampsOrg[1]:=VH_Paie.PGLibelleOrgStat1;
ChampsOrg[2]:=VH_Paie.PGLibelleOrgStat2;
ChampsOrg[3]:=VH_Paie.PGLibelleOrgStat3;
ChampsOrg[4]:=VH_Paie.PGLibelleOrgStat4;
Champslibre[1] :=VH_Paie.PgLibCombo1;
Champslibre[2] :=VH_Paie.PgLibCombo2;
Champslibre[3] :=VH_Paie.PgLibCombo3;
Champslibre[4] :=VH_Paie.PgLibCombo4;
  if H<>nil then H.Free ;
  H:=THplanning.Create(Application);
  H.Parent := PPanel ;
  H.Align:=alClient;
  H.activate:=False ;
  { Les évènements }
  H.OnDblClickSpec:=DoubleCLickSpecPlanning;
  H.OnDblClick:=DoubleCLickPlanning ;
  H.OnAvertirApplication:=HAvertirApplication;
  H.Personnalisation:=False;
  H.Legende:=True;
  { Ressources }
  H.ResChampID:='PST_CODESTAGE' ;
  { Etats }
  H.EtatChampCode:='CODE' ;
  H.EtatChampLibelle:='LIBELLE' ;
  H.EtatChampBackGroundColor:='BACKGROUNDCOLOR';
  H.EtatChampFontColor:='FONDCOLOR';
  H.EtatChampFontStyle:='FONTSTYLE';
  H.EtatChampFontSize:='FONTSIZE';
  H.EtatChampFontName:='FONTNAME';
  H.TextAlign:=taLeftJustify;
  { Items }
  H.ChampLineID    := 'PSS_CODESTAGE' ;
  H.ChampdateDebut := 'PSS_DATEDEBUT' ;
  H.ChampDateFin   := 'PSS_DATEFIN' ;
  H.ChampEtat      := 'ETAT';
  H.ChampLibelle   := 'PSS_LIBELLE';
  H.ChampHint:='PSS_LIBELLE';
  GBRecap.visible:=True;
  {Personalisation}
  H.Interval       := PiJour; // ;; //piJour
  H.CumulInterval  := pciSemaine ;
  H.IntervalDebut := PlanningFormation.DateDebFor;
  H.IntervalFin   := PlanningFormation.DateFinFor;
  H.MultiLine      := False ;
  H.GestionJoursFeriesActive := True ;
  H.ActiveSaturday := True ;
  H.ActiveSunday   := True ;
  H.ColorJoursFeries  := StringToColor('12189695');
  H.ColorOfSaturday := StringToColor('12189695');
  H.ColorOfSunday   := StringToColor('12189695');
  H.ColorBackground :=StringToColor(ColorBackground);
  H.ColorSelection:=clTeal;//StringToColor(FontColor);
  H.ActiveLigneDate := True ;
  H.ActiveLigneGroupeDate := True ;
  { Alignement des colonnes fixes }

  { Liste des champs d'entete }
  H.TokenFieldColEntete := '';
  { Liste des champs d'entete }
 // H.TokenFieldColFixed:=ListeChampRupt(';')+'PST_CODESTAGE;PST_LIBELLE';
 //PT4 - Début
 {$IFDEF EMANAGER}
 H.TokenFieldColFixed:='PST_LIBELLE';
 H.TokenSizeColFixed:= '300';
 {$ELSE}
 H.TokenFieldColFixed:='PST_CODESTAGE;PST_LIBELLE';
  if Rupture.NiveauRupt=0 then begin  H.TokenSizeColFixed:= '75;180'; H.TokenAlignColFixed := 'C;L'; end
    else
      if Rupture.NiveauRupt=1 then begin H.TokenSizeColFixed:= '150;75;180'; H.TokenAlignColFixed := 'L;C;L'; end
      else
        if Rupture.NiveauRupt=2 then begin H.TokenSizeColFixed:= '150;150;75;180'; H.TokenAlignColFixed := 'L;L;C;L'; end
        else
          if Rupture.NiveauRupt=3 then begin H.TokenSizeColFixed:= '150;150;150;75;180'; H.TokenAlignColFixed := 'L;L;L;C;L'; end
          else
            if Rupture.NiveauRupt=4 then begin H.TokenSizeColFixed:= '150;150;150;150;75;180'; H.TokenAlignColFixed := 'L;L;L;L;C;L'; end;
 {$ENDIF}
 //PT4 - Fin
  { Taille des colonnes fixes }
  H.JourneeDebut := StrToTime('08:00') ;
  H.JourneeFin   := StrToTime('18:00') ;
  H.DateFormat   := 'dd mm' ;  //
  H.DebutAM      := StrToTime('08:00') ;
  H.FinAM        := StrToTime('12:00') ;
  H.DebutAP      := StrToTime('14:00') ;
  H.FinAP        := StrToTime('18:00') ;
  H.DisplayOptionCopie:=False;
  H.DisplayOptionCreation:=False;
  H.DisplayOptionDeplacement:=False;
  H.DisplayOptionEtirer:=False;
  H.DisplayOptionLiaison:=False;
  H.DisplayOptionLier:=False;
  H.DisplayOptionModification:=False;
  H.DisplayOptionReduire:=False;
  H.DisplayOptionSuppression:=False;
  H.DisplayOptionSuppressionLiaison:=False;
  H.MoveHorizontal:=False;
  H.Autorisation:=[];
  H.ColorSelection:=H.ColorBackground ;
  
  //PT4 - Début
  {$IFDEF EMANAGER}
  CLR4.visible := False; LBLCLR4.visible := False;
  LBLCLR3.Caption := 'Session réalisée';
  LBLCLR1.Caption := 'Session ouverte';
  LBLCLR2.Caption := 'Session clôturée';
  {$ENDIF}
  //PT4 - Fin

  {Création des trois TOB }
  //PT3 - Début
  //if not ChargeTobPlaning then PostMessage(Handle, SBWM_APPSTARTUP, 0, 0)
  if not ChargeTobPlaning then
  begin
     H.Visible := False;
     PostMessage(Handle, SBWM_APPSTARTUP, 0, 0)
  end
  //PT3 - Fin
  else
    begin
     H.Visible := True; //PT3
    LoadConfigPlanning ;
    Caption:='Planning des formations du '+DateToStr(H.IntervalDebut)+' au '+DateToStr(H.IntervalFin);
    H.activate:=True ;
     BtnDetail.Click; //pt2
    end ;
end;

procedure TPlanningFormation.LoadConfigPlanning ;
var
  S,ST: String ;
  L: TStrings ;
begin
  L:=TStringList.Create;
  S:=GetSynRegKey('ConfigPlanningSessions',S,True) ;
  if S<>'' then
    begin
    L.Text:=S ;

    { Les couleurs Fond+Fériés}
    S:=L.Strings[0] ;
    H.ColorBackground  := StringToColor(ReadTokenSt(S)) ;
    H.ColorSelection   := H.ColorBackground ;
    H.ColorJoursFeries := StringToColor(ReadTokenSt(S)) ;

    { La forme graphique }
    S:=L.Strings[1] ;
    H.FormeGraphique:=AglPlanningStringToFormeGraphique(ReadTokenSt(S)) ;

    { Hauteur,Largeur, cumuldate et mode de cumul }
    S:=L.Strings[2] ;
    H.RowSizeData:=ValeurI(ReadTokenSt(S)) ;
    H.ColSizeData:=ValeurI(ReadTokenSt(S)) ;
    H.ActiveLigneGroupeDate:=ReadTokenSt(S)='1';
    H.CumulInterval:=AglPlanningStringToCumulInterval(ReadTokenSt(S)) ;
    St:=ReadTokenSt(S) ;
    if St='' then
      if (H.ColSizeData=15) AND (H.CumulInterval=pciMois) then St:='dd' else St:='dd/mm' ;
    H.DateFormat:=St ;
    end ;
  L.Free ;
end ;


function TPlanningFormation.ChargeTobPlaning: boolean;
var
Q : TQuery;
T : Tob;
DateCalculer : string;
I: Integer ;
{$IFDEF EMANAGER}NbInsc: Integer;{$ENDIF} //PT4
begin
  //result:=False;  //Pt3
  If TobItems<>nil then TobItems.Free;
  If TobEtats<>nil then TobEtats.Free;
  If TobRessources<>nil then TobRessources.Free;
//  TobRessources := TOB.Create('Les ressources',Nil,-1); //PT3
  TobEtats      := TOB.Create('Les etats',Nil,-1);
  {Codes des différents états :
  - INC : Inscriptions non clôturées
  - IC : Inscriptions clôturées
  - EFF : Session Effectuée
  - VAL : Session validée OPCA}
  T := Tob.Create('Fille etats',TobEtats,-1);
  T.AddChampSupValeur('CODE','INC',false);
  T.AddChampSupValeur('LIBELLE','Insc. non clôturées',false);
  T := Tob.Create('Fille etats',TobEtats,-1);
  T.AddChampSupValeur('CODE','IC',false);
  T.AddChampSupValeur('LIBELLE','Insc. clôturées',false);
  T := Tob.Create('Fille etats',TobEtats,-1);
  T.AddChampSupValeur('CODE','EFF',false);
  T.AddChampSupValeur('LIBELLE','Session Effectuée',false);
  T := Tob.Create('Fille etats',TobEtats,-1);
  T.AddChampSupValeur('CODE','VAL',false);
  T.AddChampSupValeur('LIBELLE','Session validée OPCA',false);

  for i:=0 to TobEtats.Detail.COunt-1 do
    Begin
    T:=TobEtats.Detail[i] ;
    if (T.getValue('CODE')='INC') then
      BackGroundColor := '10944422'
    else
      if (T.getValue('CODE')='IC') then
        BackGroundColor := '16053248'
        else
      if (T.getValue('CODE')='EFF') then
        BackGroundColor:= '10210815'
      else
        BackGroundColor:='8421631';
    T.AddChampSup('BACKGROUNDCOLOR',False);
    T.PutValue('BACKGROUNDCOLOR',BackGroundColor);
    T.AddChampSup('FONDCOLOR',False);
    T.PutValue('FONDCOLOR',FontColor);
    T.AddChampSup('FONTNAME',False);
    T.PutValue('FONTNAME',FontName);
    T.AddChampSup('FONTSTYLE',False);
    T.PutValue('FONTSTYLE',FontStyle);
    T.AddChampSup('FONTSIZE',False);
    T.PutValue('FONTSIZE',FontSize);
    End;
  Q := OpenSQL('SELECT DISTINCT PST_CODESTAGE,PST_LIBELLE FROM STAGE '+StWhTobRessource,True);
  TobRessources := Tob.Create('stages',Nil,-1);
  TobRessources.LoadDetailDB('STAGEs','','',Q,False);
  Ferme(Q);
  Q := OpenSQL('SELECT * FROM SESSIONSTAGE '+StWhTobItems,True);
  TobItems := Tob.Create('lesstages',Nil,-1);
  TobItems.LoadDetailDB('SESSIONSTAGEs','','',Q,False);
  Ferme(Q);
  For i := 0 to TobItems.Detail.Count - 1 do
  begin
//        SHeureDebut := FormatDateTime('hh:nn:ss',TobItems.Detail[i].GetValue('PSS_HEUREDEBUT'));
//        DateCalculer := FormatDateTime('dd/mm/yyyy',TobItems.Detail[i].GetValue('PSS_DATEDEBUT'));
//        DateCalculer := DateCalculer+' '+SHeureDebut;
        DateCalculer:=DateToStr(TobItems.Detail[i].GetValue('PSS_DATEDEBUT'))+' '+FloatToStrTime(8,'HH:MM:SS');
       TobItems.Detail[i].PutValue('PSS_DATEDEBUT',StrToDateTime(DateCalculer));
        DateCalculer:=DateToStr(TobItems.Detail[i].GetValue('PSS_DATEFIN'))+' '+FloatToStrTime(17,'HH:MM:SS');
        TobItems.Detail[i].PutValue('PSS_DATEFIN',StrToDateTime(DateCalculer));
        
        //PT4 - Début
        {$IFDEF EMANAGER}
        If TobItems.Detail[i].GetValue('PSS_DATEFIN') < Date Then
        	TobItems.Detail[i].AddChampSupValeur('ETAT','EFF')
        Else
        Begin
	        Q := OpenSQL ('SELECT COUNT(*) AS NBINSC FROM FORMATIONS WHERE PFO_CODESTAGE="'+TobItems.Detail[i].GetValue('PSS_CODESTAGE')+'" AND PFO_ORDRE="'+IntToStr(TobItems.Detail[i].GetValue('PSS_ORDRE'))+'" AND PFO_MILLESIME="'+TobItems.Detail[i].GetValue('PSS_MILLESIME')+'"', True);
	        If Not Q.EOF Then
	        	NbInsc := Q.FindField('NBINSC').AsInteger
	        Else
	        	NbInsc := 0;
	        Ferme(Q);
			If TobItems.Detail[i].Getvalue('PSS_NBRESTAGPREV') > NbInsc Then
				TobItems.Detail[i].AddChampSupValeur('ETAT','INC')
			Else
				TobItems.Detail[i].AddChampSupValeur('ETAT','IC');
        End;
        {$ELSE}
        TobItems.Detail[i].AddChampSupValeur('ETAT','INC');
        If TobItems.Detail[i].Getvalue('PSS_CLOTUREINSC') = 'X' then TobItems.Detail[i].PutValue('ETAT','IC');
        If TobItems.Detail[i].Getvalue('PSS_EFFECTUE') = 'X' then TobItems.Detail[i].PutValue('ETAT','EFF');
        If TobItems.Detail[i].Getvalue('PSS_VALIDORG') = 'X' then TobItems.Detail[i].PutValue('ETAT','VAL');
        {$ENDIF}
        //PT4 - Fin
   end;
//  MiseEnFormeTobRupt('RES',Tob_Sal);
(*
  TProgress.Value:=TProgress.Value+1;
  TProgress.SubText:='Affectation du PlanningFormation..';
*)
  H.TobRes:=TobRessources ;
  H.TobEtats:=TobEtats;
  H.TobItems:=TobItems;
  //PT1 - Début
  If TobRessources.FillesCount(0) = 0 Then
  Begin
       Result := False;
       If TobItems<>nil then TobItems.Free;
       If TobEtats<>nil then TobEtats.Free;
       If TobRessources<>nil then TobRessources.Free;
       PGIBox('La sélection ne renvoie aucun résultat.');
  End
  else
  //PT1 - Fin
(*
  TProgress.free;
*)
//if Tob_Sal<>nil then Tob_Sal.free;
  Result:=True;
end;

{procedure TPlanningFormation.MiseEnFormeTobRupt(TypeTob : string ; Tob_Stages : Tob);
var I : Integer;
begin
if TypeTob='RES' then
  for I:=0 to Tob_Stages.Detail.Count-1 do
     AddRessTob(Tob_Stages.Detail[i]);
end;}

{Function TPlanningFormation.AddRessTob(Tob_Stages : Tob) : Boolean;
var RessSal : String;
begin
result:=False;
While Result=False do
  Begin
  RessSal:=Tob_Stages.GetValue('PST_CODESTAGE');
  result:=CreateEnr('PST_CODESTAGE',Tob_Stages);   //Niv 0
  end;
End;}

{Function TPlanningFormation.CreateEnr(Champ: string; Tob_Stages: Tob) : Boolean;
Var T1 : Tob;
begin
result:=False;
T1:=Tob.Create('Les ressources',TobRessources,-1);
With T1 do
  Begin
  AddChampSupValeur('PST_CODESTAGE','');
  AddChampSupValeur('PST_LIBELLE','');
  if Champ='PST_CODESTAGE' then
     Begin
     PutValue('PST_CODESTAGE',Tob_Stages.Getvalue('PST_CODESTAGE'));
     PutValue('PST_LIBELLE',Tob_Stages.GetValue('PST_LIBELLE'));
     Result:=True;
     end;
  end;
end; }


procedure TPlanningFormation.FormClose(Sender: TObject; var Action: TCloseAction);
begin
       If H <> Nil Then H.Free ;
       If TobItems<>nil then TobItems.Free;
       If TobEtats<>nil then TobEtats.Free;
       If TobRessources<>nil then TobRessources.Free;
end;


procedure TPlanningFormation.BFermeClick(Sender: TObject);
begin
PlanningFormation.Close;
end;

{function TPlanningFormation.ListeChampRupt(Sep : string) : string;
Var Prefixe : string;
begin
  if Sep=';' then Prefixe:='LIB_' else Prefixe:='';
  if Rupture.ChampsRupt[1]<>'' then Result:=Prefixe+Rupture.ChampsRupt[1]+Sep;
  if Rupture.ChampsRupt[2]<>'' then Result:=Result+Prefixe+Rupture.ChampsRupt[2]+Sep;
  if Rupture.ChampsRupt[3]<>'' then Result:=Result+Prefixe+Rupture.ChampsRupt[3]+Sep;
  if Rupture.ChampsRupt[4]<>'' then Result:=Result+Prefixe+Rupture.ChampsRupt[4]+Sep;
end;}



procedure TPlanningFormation.HAvertirApplication(Sender: TObject; FromItem,
  ToItem: TOB; Actions: THPlanningAction);
Var
Stage,Millesime : string;
NumOrdre : Integer;
begin
IF Actions=paclickLeft THEN
  begin
  IF FromItem<>Nil THEN
    begin
    Stage:=FromItem.GetValue('PSS_CODESTAGE');
    Millesime:=FromItem.GetValue('PSS_MILLESIME');
    NumOrdre:=FromItem.GetValue('PSS_ORDRE');
    RefrehRecapPlanning(Stage,Millesime,NumOrdre,FromItem);
    end ;
  end ;
end;

procedure TPlanningFormation.DoubleCLickPlanning(Sender: TObject);
var
  T: TOB ;
  Stage,Millesime,ret,St : String;
  Q : TQuery;
  NumOrdre : Integer;
  DDSession,DFSession : TDateTime;
begin
T:=H.GetCurItem;
if T<>Nil then
  begin
  Stage := T.GetValue('PSS_CODESTAGE');
  NumOrdre :=T.GetValue('PSS_ORDRE');
  Millesime := T.GetValue('PSS_MILLESIME');
  St := Stage+';'+IntToStr(NumOrdre)+';'+Millesime;
  DDSession := T.GetValue('PSS_DATEDEBUT');
  DFSession := T.GetValue('PSS_DATEFIN');
  //PT1 - Début
  {$IFNDEF EMANAGER}
  If ((DDSession >= DDMillesimeEC) and (DDSession <=DFMillesimeEC)) or ((DFSession >= DDMillesimeEC) and (DFSession <=DFMillesimeEC)) then
        Ret := AglLanceFiche ('PAY','SESSIONSTAGE', '', St, 'SAISIE'+';;;'+MillesimeEC)
  else Ret := AglLanceFiche ('PAY','SESSIONSTAGE', '', St, 'SAISIE'+';;;0000');
  {$ELSE}
  If ((DDSession >= DDMillesimeEC) and (DDSession <=DFMillesimeEC)) or ((DFSession >= DDMillesimeEC) and (DFSession <=DFMillesimeEC)) then
        Ret := AglLanceFiche ('PAY','EM_SESSIONSTAGE', '', St, 'CONSULTATION'+';;;'+MillesimeEC)
  else Ret := AglLanceFiche ('PAY','EM_SESSIONSTAGE', '', St, 'CONSULTATION'+';;;0000');
  {$ENDIF}
  //PT1 - Fin
  If Ret = 'SUPPR' then
  begin
       H.DeleteItem(T,False);
       T.free; T:=Nil;
  end;
  If Ret = 'MODIF' then
  begin
       Q := OpenSQL('SELECT PSS_LIBELLE,PSS_DATEDEBUT,PSS_DATEFIN,PSS_EFFECTUE,PSS_CLOTUREINSC,PSS_VALIDORG,PSS_LIEUFORM,PSS_NATUREFORM FROM SESSIONSTAGE '+
       'WHERE PSS_CODESTAGE="'+T.GetValue('PSS_CODESTAGE')+'" AND PSS_MILLESIME="'+T.GetValue('PSS_MILLESIME')+'" AND '+
       'PSS_ORDRE='+IntToStr(T.GetValue('PSS_ORDRE')),True);
       If Not Q.Eof then
       begin
         T.PutValue('ETAT','INC');
         If Q.FindField('PSS_CLOTUREINSC').AsString = 'X' then T.PutValue('ETAT','IC');
         If Q.FindField('PSS_EFFECTUE').AsString = 'X' then T.PutValue('ETAT','EFF');
         If Q.FindField('PSS_VALIDORG').AsString = 'X' then T.PutValue('ETAT','VAL');
         T.PutValue('PSS_LIBELLE',Q.FindField('PSS_LIBELLE').AsString);
         T.PutValue('PSS_LIEUFORM',Q.FindField('PSS_LIEUFORM').AsString);
         T.PutValue('PSS_NATUREFORM',Q.FindField('PSS_NATUREFORM').AsString);
         T.PutValue('PSS_DATEDEBUT',Q.FindField('PSS_DATEDEBUT').AsDateTime);
         T.PutValue('PSS_DATEFIN',Q.FindField('PSS_DATEFIN').AsDateTime);
       end;
       Ferme(Q);
       RefrehRecapPlanning(T.GetValue('PSS_CODESTAGE'),T.GetValue('PSS_MILLESIME'),T.GetValue('PSS_ORDRE'),Nil);
  end;
  end;
end;

procedure TPlanningFormation.DoubleCLickSpecPlanning(ACol, ARow: INTEGER;TypeCellule: TPlanningTypeCellule; T: TOB);
begin
//  CASE TypeCellule OF
//    ptcRessource       : if T<>nil then
//                            RefrehRecapPlanning(T.GetValue('PSS_CODESTAGE'),T.GetValue('PSS_MILLESIME'),T.GetValue('PSS_ORDRE'),nil);
//  END ;
end;

procedure TPlanningFormation.RefrehRecapPlanning(Stage,Millesime : string;Numordre:Integer; TItem : Tob);
var Q : TQuery;
    NbInsc : Integer;//PT2
begin
        Q:= OpenSQL('SELECT PSS_LIBELLE,PSS_NATUREFORM,PSS_DATEDEBUT,PSS_DATEFIN,PSS_LIEUFORM,PSS_NBRESTAGPREV FROM SESSIONSTAGE'+ //PT2
        ' WHERE PSS_CODESTAGE="'+Stage+'" AND PSS_MILLESIME="'+Millesime+'" AND PSS_ORDRE='+IntToStr(NumOrdre),True);
        If Not Q.Eof then
        begin
                HSTAGE.Caption := RechDom('PGSTAGEFORM',Stage,False);
                HSESSION.Caption := Q.FindField('PSS_LIBELLE').AsString;
                HNATURE.Caption := RechDom('PGNATUREFORM',Q.FindField('PSS_NATUREFORM').AsString,False);
                HLIEU.Caption := RechDom('PGLIEUFORMATION',Q.FindField('PSS_LIEUFORM').AsString,False); //PT4
                HDATEDEB.Caption := FormatDateTime('dd/mm/yyyy',Q.FindField('PSS_DATEDEBUT').AsDateTime);
                HDATEFIN.Caption := FormatDateTime('dd/mm/yyyy',Q.FindField('PSS_DATEFIN').AsDateTime);
                NbInsc := Q.FindField('PSS_NBRESTAGPREV').AsInteger;  //PT2
        end
        else
        begin
                HSTAGE.Caption := '';
                HSESSION.Caption := '';
                HNATURE.Caption := '';
                HLIEU.Caption := '';
                HDATEDEB.Caption := '';
                HDATEFIN.Caption := '';
                NbInsc := 0;  //PT2
        end;
        Ferme(Q);
        //PT2 - Début
        Q := OpenSQL ('SELECT COUNT(*) AS NBINSC FROM FORMATIONS WHERE PFO_CODESTAGE="'+Stage+'" AND PFO_ORDRE="'+IntToStr(NumOrdre)+'" AND PFO_MILLESIME="'+Millesime+'"', True);
        If Not Q.EOF Then
        Begin
                HNBDISPO.Caption := IntToStr(NbInsc - Q.FindField('NBINSC').AsInteger);
                HNBINSC.Caption := IntToStr(Q.FindField('NBINSC').AsInteger);
        End
        Else
        Begin
                HNBDISPO.Caption := '';
                HNBINSC.Caption := '';
        End;
        Ferme(Q);
        //PT2 - Fin
end;

procedure TPlanningFormation.BtnDetailClick(Sender: TObject);
begin
PRecap.visible:=BtnDetail.Down;
end;

procedure TPlanningFormation.BImprimerClick(Sender: TObject);   
var
  S: String ;
begin
  H.TypeEtat:='E' ;
  H.NatureEtat:='PFO' ;
  H.CodeEtat:='PLG' ;
  S:='DATEDEB='+FormatDateTime('dd/mm/yyyy',H.IntervalDebut)+'^E'+IntToStr(Byte(otDate)) ;
  S:=S+'`DATEFIN='+FormatDateTime('dd/mm/yyyy',H.IntervalFin)+'^E'+IntToStr(Byte(otDate)) ;
  H.CriteresToPrint:=S ;
  H.Print ;
end;

procedure TPlanningFormation.TExcelClick(Sender: TObject);
begin
  with TSaveDialog.Create(Self) do
  begin
    Filter := 'Microsoft Excel (*.xls)|*.XLS';
    if Execute then H.ExportToExcel(True, FileName);
    Free;
  end;
end;


end.

