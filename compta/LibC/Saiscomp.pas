{***********UNITE*************************************************
Auteur  ...... : Je ne m'en rappelle plus
Créé le ...... : 01/06/2004
Modifié le ... : 27/10/2005
Description .. : Récupération de l'ancienne gestion du blocnote car suite
Suite ........ : aux passage en Web Access, le traitements des blocnotes
Suite ........ : ne fonctionnait pas.
Suite ........ : ==> 2 fonctionnements : un Web Access nouvellement
Suite ........ : écrit, et un 2/3 comme avant. (CA - 01/06/2004)
Suite ........ :
Suite ........ : JP 08/08/05 : FQ 14919 : Appel à la gestion d'affaires
Suite ........ :
Suite ........ : JP 18/10/05 : FQ 16899 : On n'exécute plus
Suite ........ : VerifObligatoire si on est en CutOff
Suite ........ :
Suite ........ :
Suite ........ : JP 27/10/05 : A la demande de JL Decosse, remplacement
Suite ........ : des BitBtns par des SpeedButtons
Mots clefs ... :
*****************************************************************}
unit SaisComp;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mask, StdCtrls, Hctrls, Buttons, SaisUtil , ExtCtrls, HEnt1,
  UTofAffaire_Mul_Comp,  // AGLLanceFiche ('AFF','AFFAIRE_MUL_COMPT','  unite manquante
{$IFDEF EAGLCLIENT}
	MaineAGL,
{$ELSE}
  DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  fe_main,
  // Scenario, BVE 21.05.07 FQ 19684
{$ENDIF}
{$IFNDEF CCS3}
  LibChpLi,
{$ENDIF}
{$IFDEF VER150}
   Variants,
{$ENDIF}
  UTOB ,
  hmsgbox, ComCtrls, HRichEdt, HSysMenu, HRichOLE, Hcompte, ParamSoc, UlibEcriture, uLibExercice,
  HTB97;

Type R_COMP = RECORD
              StComporte,StLibre : String ;
              Conso,Attributs : boolean ;
              MemoComp        : HTStrings ;
              Origine         : integer ;
              DateC           : TDateTime ;
              CutOffPer       : string ;
              CutOffEchue     : string ;
              TOBCompl        : TOB ;
              AvecCalcul      : boolean ;   
              END ;

Function SaisieComplement ( OBM : TOBM ; QuelEcr : TTypeEcr ; Action : TActionFiche ;
                            Var ModBN : boolean ; RC : R_COMP ; LibreAnaObli : boolean = False ; FocusBN : boolean = False ) : boolean ;

//Function SaisieComplement ( OBM : TOBM ; QuelEcr : TTypeEcr ; Action : TActionFiche ;
//                           Var ModBN : boolean ; RC : R_COMP ) : boolean ;
Function ComporteLigne ( Scenario : TOBM ; General : String17 ; Var StComporte,StLibre : String ) : integer ;


type
  TFSaisComp = class(TForm)
    TOTQ1: TEdit;
    TOTQ2: TEdit;
    HC: THMsgBox;
    Panel1: TPanel;
    HINFO: TLabel;
    HUser: TLabel;
    HMTrad: THSystemMenu;
    HM: THMsgBox;
    HCptPerso: THCpteEdit;
    PRef: TPanel;
    GInt: TGroupBox;
    H_REFINTERNE: THLabel;
    H_LIBELLE: THLabel;
    _REFINTERNE: TEdit;
    _LIBELLE: TEdit;
    Pmenfou: TPanel;
    Pages: TPageControl;
    TS1: TTabSheet;
    GBloc: TGroupBox;
    _BLOCNOTE: THRichEditOLE;
    TS2: TTabSheet;
    Bevel3: TBevel;
    Bevel2: TBevel;
    Bevel1: TBevel;
    H_LIBRETEXTE0: THLabel;
    H_LIBRETEXTE2: THLabel;
    H_LIBRETEXTE4: THLabel;
    H_LIBRETEXTE6: THLabel;
    H_LIBRETEXTE8: THLabel;
    H_LIBRETEXTE1: THLabel;
    H_LIBRETEXTE3: THLabel;
    H_LIBRETEXTE5: THLabel;
    H_LIBRETEXTE7: THLabel;
    H_LIBRETEXTE9: THLabel;
    H_LIBREMONTANT0: THLabel;
    H_LIBREMONTANT2: THLabel;
    H_LIBREMONTANT1: THLabel;
    H_LIBREMONTANT3: THLabel;
    H_LIBREDATE: THLabel;
    H_TABLE2: THLabel;
    H_TABLE0: THLabel;
    H_TABLE1: THLabel;
    H_TABLE3: THLabel;
    _LIBRETEXTE0: TEdit;
    _LIBRETEXTE2: TEdit;
    _LIBRETEXTE4: TEdit;
    _LIBRETEXTE6: TEdit;
    _LIBRETEXTE8: TEdit;
    _LIBRETEXTE1: TEdit;
    _LIBRETEXTE3: TEdit;
    _LIBRETEXTE5: TEdit;
    _LIBRETEXTE7: TEdit;
    _LIBRETEXTE9: TEdit;
    _LIBREMONTANT1: THNumEdit;
    _LIBREMONTANT3: THNumEdit;
    _LIBREMONTANT0: THNumEdit;
    _LIBREMONTANT2: THNumEdit;
    _LIBREDATE: THCritMaskEdit;        {FP FQ15916}
    _LIBREBOOL0: TCheckBox;
    _LIBREBOOL1: TCheckBox;
    _TABLE0: THCpteEdit;
    _TABLE2: THCpteEdit;
    _TABLE1: THCpteEdit;
    _TABLE3: THCpteEdit;
    PCUTOFF: TPanel;
    GroupBox1: TGroupBox;
    HLabel1: THLabel;
    CUTOFFDEB: THCritMaskEdit;
    HLabel2: THLabel;
    CUTOFFFIN: THCritMaskEdit;
    PInfo: TPanel;
    GExt: TGroupBox;
    H_REFEXTERNE: THLabel;
    H_DATEREFEXTERNE: THLabel;
    _REFEXTERNE: TEdit;
    _DATEREFEXTERNE: THCritMaskEdit;    {FP FQ15916}
    GComp: TGroupBox;
    H_REFLIBRE: THLabel;
    H_AFFAIRE: THLabel;
    H_CONSO: THLabel;
    H_NATURETRESO: THLabel;
    _REFLIBRE: TEdit;
    _AFFAIRE: TEdit;
    _CONSO: TEdit;
    _NATURETRESO: THValComboBox;
    GQte: TGroupBox;
    H_QTE1: THLabel;
    H_QTE2: THLabel;
    H_QUALIFQTE1: THLabel;
    H_QUALIFQTE2: THLabel;
    H_POURCENTQTE1: THLabel;
    H_POURCENTQTE2: THLabel;
    _QTE1: THNumEdit;
    _QTE2: THNumEdit;
    _QUALIFQTE1: THValComboBox;
    _QUALIFQTE2: THValComboBox;
    _POURCENTQTE1: THNumEdit;
    _POURCENTQTE2: THNumEdit;
    CUTOFFPERIODE: THValComboBox;
    HLabel3: THLabel;
    BDetail: TToolbarButton97;
    BReduire: TToolbarButton97;
    BValide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure BValideClick(Sender: TObject);
    procedure _QTE1Exit(Sender: TObject);
    procedure _QTE2Exit(Sender: TObject);
    procedure _POURCENTQTE1Exit(Sender: TObject);
    procedure _POURCENTQTE2Exit(Sender: TObject);
    procedure _BLOCNOTEChange(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BAbandonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure _REFEXTERNEChange(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BAideClick(Sender: TObject);
    procedure VerifLibre(Sender: TObject);
    procedure DBCPerso(Sender: TObject);
    procedure _AFFAIREKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure CUTOFFFINExit(Sender: TObject);
    procedure BDetailClick(Sender: TObject);
    procedure BReduireClick(Sender: TObject);
    procedure CUTOFFDEBExit(Sender: TObject);
    procedure CUTOFFPERIODEChange(Sender: TObject);
    procedure _AFFAIREDblClick(Sender: TObject);
  private
    OB : TOBM ;
    Pf : String ;
    FBoCutOff : boolean ;
    procedure EtudieQtes ;
    procedure EtudieComporte ;
    procedure InitTitre ;
    Procedure GoEnable(St : String) ;
    procedure EnableCeQuiFaut ;
    procedure InitZones ;
    Function  MakeInfoUser ( OB : TOB ; Tip : TTypeEcr ) : String ;
    procedure RendEna ( C : TControl ) ;
    procedure ChangeLesMask ;
    Function  ZoneVide ( C : TComponent ) : boolean ;
    Function  VerifObligatoire : boolean ;
    Function  ZoneObli ( Nam : String ) : boolean ;
    Procedure ValeurDefModif ( C : TComponent ) ;
    Function  ConvertNom ( Nam : String ) : String ;
    procedure ValideLesComps ;
    procedure GereAffaire;
    procedure RemplitLesZones ;
    procedure CalculDateCutOff ;
  public
    OKV,GeneModif : boolean ;
    QuelEcr       : TTypeEcr ;
    Action        : TActionFiche ;
    FocusBN,ModifBN,LibreAnaObli : boolean ;
    RC            : R_COMP ;
  end;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcGen,
  {$ENDIF MODENT1}
  Ent1;


Procedure ValeurAttributComp(Li : HTStrings ; Radical,Champ : String ; Var Oblg,Modif,Valdef,TexteLibre : String ) ;
Var i,j : Integer ;
    St : String ;
BEGIN
Oblg:='-' ; Modif:='X' ; Valdef:='' ; TexteLibre:='' ;
if (Champ='') or (Li=Nil) or (Li.Count<=0) then Exit ;
for i:=0 to Li.Count-1 do
   BEGIN
   St:=Li.Strings[i] ;
   if Pos(Radical,St)<=0 then Continue ;
   if Pos(Champ,St)>0 then
      BEGIN
      for j:=1 to 2 do ReadTokenSt(St) ;
      Oblg:=ReadTokenSt(St) ; Modif:=ReadTokenSt(St) ; ValDef:=ReadTokenSt(St) ;
      TexteLibre:=ReadTokenSt(St) ;
      Break ;
      END ;
   END ;
END ;

Function SaisieComplement ( OBM : TOBM ; QuelEcr : TTypeEcr ; Action : TActionFiche ;
                            Var ModBN : boolean ; RC : R_COMP ;  LibreAnaObli : boolean = False ; FocusBN : boolean = False ) : boolean ;
Var X : TFSaisComp ;
BEGIN
Result:=False ;
if ((Pos('X',RC.StComporte)<=0) and (Pos('X',RC.StLibre)<=0)) then Exit ;
X:=TFSaisComp.Create(Application) ;
 Try
  X.OB:=OBM ; X.QuelEcr:=QuelEcr ; X.Action:=Action ;
  X.ModifBN:=False ; X.OkV:=False ; X.RC:=RC ;
  X.LibreAnaObli:=LibreAnaObli ; X.FocusBN:=FocusBN ;
  Result:=((X.ShowModal=mrOK) or (X.OkV)) ; if Result then ModBN:=X.ModifBN ;
 Finally
  X.Free ;
 End ;
Screen.Cursor:=SyncrDefault ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 30/08/2005
Modifié le ... : 12/05/2006
Description .. : - FB 15874 - on a pas acces a la date de fin si la periode est
Suite ........ : diff de aucun
Suite ........ : - LG - 12/05/2006 - on supprime la case periode en 
Suite ........ : consultation
Mots clefs ... : 
*****************************************************************}
procedure TFSaisComp.EtudieComporte ;
Var i,ip : integer ;
    C,PremC : TControl ;
    Exo : TExoDate;
BEGIN

if FBoCutOff then
 BreduireClick(nil)
  else
    begin
    PCUTOFF.Visible     := RC.TOBCompl <> nil ;
    // YMO 05/12/2006 NORME NF:E_REFINTERNE Disabled sous conditions(exercice clos ou ecriture validée)
    if (Pf='E') and not QuelDateDeExo(OB.GetValue('E_EXERCICE'),Exo) then
      begin
      If ((OB.GetValue('E_VALIDE')='X') Or (Exo.EtatCpta='CLO')) then
        RC.StComporte[1]:='-';
      end;
    end ;

 CUTOFFPERIODE.OnChange :=  nil ;

 if not RC.AvecCalcul then
  begin
   CUTOFFPERIODE.visible := false ;
   HLabel3.visible       := false ;
  end
   else
    begin
     if RC.CutOffPer <> '' then
      CUTOFFPERIODE.Value := RC.CutOffPer ;
     CUTOFFFIN.Enabled      := ( CUTOFFPERIODE.Text = 'Aucun' ) or ( RC.CutOffEchue = 'X' ) ;
     CUTOFFDEB.Enabled      := ( CUTOFFPERIODE.Text = 'Aucun' ) or ( RC.CutOffEchue = '-' ) ;
    end ;
    
 CUTOFFPERIODE.OnChange := CUTOFFPERIODEChange  ;

if FBoCutOff then exit ;

if Pos('X',Copy(RC.StComporte,3,200))<=0 then Pages.Pages[0].TabVisible:=False ;
if Pos('X',RC.StLibre)<=0 then Pages.Pages[1].TabVisible:=False ;
PremC:=Nil ; ip:=0 ;
for i:=0 to ComponentCount-1 do
    BEGIN
    C:=TControl(Components[i]) ;
    if C is TPanel then Continue ; if C is TBitBtn then Continue ;
    if ((C.Tag>0) and (C.Tag<100)) then
       BEGIN
       C.Enabled:=((C.Enabled) and (RC.StComporte[C.Tag]='X')) ;
       if ((Not C.Enabled) and (Action<>taConsult)) then TEdit(C).Color:=clBtnFace ;
       if ((PremC=Nil) and (C.Enabled)) then BEGIN PremC:=C ; ip:=0 ; END ;
       END ;
    if C.Tag>100 then
       BEGIN
       if RC.StLibre<>'' then C.Enabled:=((C.Enabled) and (RC.StLibre[C.Tag-100]='X'))
                         else C.Enabled:=False ;
       if ((Not C.Enabled) and (Action<>taConsult)) then if C is TCustomEdit then TEdit(C).Color:=clBtnFace ;
       if ((PremC=Nil) and (C.Enabled)) then BEGIN PremC:=C ; ip:=1 ; END ;
       END ;
    END ;
if ((ip=0) and (Pages.Pages[0].TabVisible)) then Pages.ActivePage:=TS1 else
   if Pages.Pages[1].TabVisible then Pages.ActivePage:=TS2 ;
if PremC<>Nil then if TWinControl(PremC).CanFocus then ActiveControl:=TWinControl(PremC) ;
END ;

procedure TFSaisComp.EtudieQtes ;
BEGIN
if Not (QuelEcr in [EcrGen,EcrBud]) then
   BEGIN
   if _QUALIFQTE1.Value='' then _QUALIFQTE1.Value:=OB.GetValue(Pf+'_QUALIFECRQTE1') ;
   if _QUALIFQTE2.Value='' then _QUALIFQTE2.Value:=OB.GetValue(Pf+'_QUALIFECRQTE2') ;
   END ;
if _QUALIFQTE1.Value='' then _QUALIFQTE1.Value:='...' ;
if _QUALIFQTE2.Value='' then _QUALIFQTE2.Value:='...' ;
// Aucun recalcul de l'affichage des qtés // MODIF Fiche 12718
{
if (Action = taCreat) and Not (QuelEcr in [EcrGen,EcrBud]) then // MODIF Fiche 12718
   BEGIN
   if ((Valeur(TOTQ1.Text)<>0) and (Valeur(_QTE1.Text)=0) and (Valeur(_POURCENTQTE1.Text)=0))
      then _POURCENTQTE1.Text:=StrfMontant(OB.GetValue(Pf+'_POURCENTAGE'),15,ADecimP,'',True) ;
   if ((Valeur(TOTQ2.Text)<>0) and (Valeur(_QTE2.Text)=0) and (Valeur(_POURCENTQTE2.Text)=0))
      then _POURCENTQTE2.Text:=StrfMontant(OB.GetValue(Pf+'_POURCENTAGE'),15,ADecimP,'',True) ;
   _POURCENTQTE1Exit(Nil) ; _POURCENTQTE2Exit(Nil) ;
   END ;
}
END ;

procedure TFSaisComp.RendEna ( C : TControl ) ;
BEGIN
C.Tag:=-1 ; C.Enabled:=False ;
END ;

Function TFSaisComp.MakeInfoUser ( OB : TOB ; Tip : TTypeEcr ) : String ;
BEGIN
Result:='' ;
Case Tip of
   EcrGen : Result:=HUser.Caption+' '+RechDom('ttUtilisateur',OB.GetValue('E_UTILISATEUR'),False)+'   '+FormatDateTime('dd/mm/yy hh:nn',OB.GetValue('E_DATEMODIF')) ;
   EcrBud : Result:=HUser.Caption+' '+RechDom('ttUtilisateur',OB.GetValue('BE_UTILISATEUR'),False)+'   '+FormatDateTime('dd/mm/yy hh:nn',OB.GetValue('BE_DATEMODIF')) ;
   EcrAna : Result:=HUser.Caption+' '+RechDom('ttUtilisateur',OB.GetValue('Y_UTILISATEUR'),False)+'   '+FormatDateTime('dd/mm/yy hh:nn',OB.GetValue('Y_DATEMODIF')) ;
   END ;
END ;

procedure TFSaisComp.ChangeLesMask ;
BEGIN
ChangeMask(_LIBREMONTANT0,V_PGI.OkDecV,'') ; ChangeMask(_LIBREMONTANT2,V_PGI.OkDecV,'') ;
ChangeMask(_LIBREMONTANT2,V_PGI.OkDecV,'') ; ChangeMask(_LIBREMONTANT3,V_PGI.OkDecV,'') ;
ChangeMask(_QTE1,V_PGI.OkDecQ,'') ; ChangeMask(_QTE2,V_PGI.OkDecQ,'') ;
END ;

procedure TFSaisComp.InitTitre ;
BEGIN
Case QuelEcr of
   EcrGen : BEGIN
            if OB.GetValue('E_GENERAL')<>'' then
               BEGIN
               HInfo.Caption:=HInfo.Caption+IntToStr(OB.GetValue('E_NUMLIGNE'))+'   '+OB.GetValue('E_GENERAL')+'   '+OB.GetValue('E_REFINTERNE') ;
               HUser.Caption:=MakeInfoUser(OB,EcrGen) ;
               END else
               BEGIN
               HInfo.Caption:=HC.Mess[0] ;
               HUser.Caption:=HUser.Caption+' '+RechDom('ttUtilisateur',V_PGI.User,False) ;
               END ;
            Pf:='E' ;
            END ;
   EcrBud : BEGIN
            if OB.GetValue('BE_BUDGENE')<>'' then HInfo.Caption:=HInfo.Caption+'   '+OB.GetValue('BE_BUDGENE')+'   '+OB.GetValue('BE_REFINTERNE') ;
            HUser.Caption:=MakeInfoUser(OB,EcrBud) ;
            Pf:='BE' ;
            END ;
   EcrAna : BEGIN
            HInfo.Caption:=HInfo.Caption+IntToStr(OB.GetValue('Y_NUMVENTIL'))+'   '+OB.GetValue('Y_SECTION')+'   '+OB.GetValue('Y_GENERAL') ;
            HUser.Caption:=MakeInfoUser(OB,EcrAna) ;
            Pf:='Y' ;
            END ;
   END ;
END ;

Procedure TFSaisComp.GoEnable(St : String) ;
Var C : TComponent ;
    CC : TControl ;
BEGIN
C:=Self.FindComponent(St) ;
If C<>NIL Then BEGIN CC:=TControl(C) ; CC.Enabled:=FALSE ; END ;
END ;

procedure TFSaisComp.EnableCeQuiFaut ;

BEGIN
GoEnable('_LIBRETEXTE0') ; GoEnable('H_LIBRETEXTE0') ;
GoEnable('_LIBRETEXTE1') ; GoEnable('H_LIBRETEXTE1') ;
GoEnable('_LIBRETEXTE2') ; GoEnable('H_LIBRETEXTE2') ;
GoEnable('_LIBRETEXTE3') ; GoEnable('H_LIBRETEXTE3') ;
GoEnable('_LIBRETEXTE4') ; GoEnable('H_LIBRETEXTE4') ;
GoEnable('_LIBRETEXTE5') ; GoEnable('H_LIBRETEXTE5') ;
GoEnable('_LIBRETEXTE6') ; GoEnable('H_LIBRETEXTE6') ;
GoEnable('_LIBRETEXTE7') ; GoEnable('H_LIBRETEXTE7') ;
GoEnable('_LIBRETEXTE8') ; GoEnable('H_LIBRETEXTE8') ;
GoEnable('_LIBRETEXTE9') ; GoEnable('H_LIBRETEXTE9') ;
If EstSerie(S3) Or EstSerie(S5) Then
  BEGIN
  GoEnable('_TABLE3') ; GoEnable('H_TABLE3') ;
  GoEnable('_TABLE2') ; GoEnable('H_TABLE2') ;
  END ;
If EstSerie(S3)  Then
  BEGIN
  GoEnable('_TABLE1') ; GoEnable('H_TABLE1') ;
  END ;
GoEnable('_LIBREMONTANT0') ; GoEnable('H_LIBREMONTANT0') ;
GoEnable('_LIBREMONTANT1') ; GoEnable('H_LIBREMONTANT1') ;
GoEnable('_LIBREMONTANT2') ; GoEnable('H_LIBREMONTANT2') ;
GoEnable('_LIBREMONTANT3') ; GoEnable('H_LIBREMONTANT3') ;
GoEnable('_LIBREBOOL0') ; GoEnable('_LIBREBOOL1') ;
GoEnable('_LIBREDATE') ;  

END ;


procedure TFSaisComp.InitZones ;
BEGIN
if Not RC.Conso then BEGIN _CONSO.Enabled:=False ; _CONSO.Color:=clBtnFace ; END ;
Case QuelEcr of
   EcrGen : BEGIN
            _POURCENTQTE1.Visible:=False ; _POURCENTQTE2.Visible:=False ;
            _POURCENTQTE1.Tag:=-1 ; _POURCENTQTE2.Tag:=-1 ;
            H_POURCENTQTE1.Visible:=False ; H_POURCENTQTE2.Visible:=False ;
            TOTQ1.Text:='' ; TOTQ2.Text:='' ;
            _TABLE0.ZoomTable:=tzNatEcrE0 ; _TABLE1.ZoomTable:=tzNatEcrE1 ;
            _TABLE2.ZoomTable:=tzNatEcrE2 ; _TABLE3.ZoomTable:=tzNatEcrE3 ;
            END ;
   EcrBud : BEGIN
            _POURCENTQTE1.Visible:=False ; _POURCENTQTE2.Visible:=False ;
            _POURCENTQTE1.Tag:=-1 ; _POURCENTQTE2.Tag:=-1 ;
            _CONSO.Visible:=False ; _CONSO.Tag:=-1 ;
            H_POURCENTQTE1.Visible:=False ; H_POURCENTQTE2.Visible:=False ;
            RendEna(_REFLIBRE) ; RendEna(_AFFAIRE) ; RendEna(_REFEXTERNE) ; RendEna(_DATEREFEXTERNE) ;
            TOTQ1.Text:='' ; TOTQ2.Text:='' ;
            _TABLE0.ZoomTable:=tzNatEcrU0 ; _TABLE1.ZoomTable:=tzNatEcrU1 ;
            _TABLE2.ZoomTable:=tzNatEcrU2 ; _TABLE3.ZoomTable:=tzNatEcrU3 ;
            _NATURETRESO.Visible:=False ; RendEna(_NATURETRESO) ;
            END ;
   EcrAna : BEGIN
// JLD 23/02/99            _REFLIBRE.Visible:=False ; _REFLIBRE.Tag:=-1 ; H_REFLIBRE.Visible:=False ;
            TOTQ1.Text:=StrfMontant(OB.GetValue('Y_TotalQte1'),15,4,'',True) ;
            TOTQ2.Text:=StrfMontant(OB.GetValue('Y_TotalQte2'),15,4,'',True) ;
            _TABLE0.ZoomTable:=tzNatEcrA0 ; _TABLE1.ZoomTable:=tzNatEcrA1 ;
            _TABLE2.ZoomTable:=tzNatEcrA2 ; _TABLE3.ZoomTable:=tzNatEcrA3 ;
            if EstSerie(S5) then GQte.Visible:=False ;
            _NATURETRESO.Visible:=False ; RendEna(_NATURETRESO) ;
            END ;
   END ;
(*if ((EstSerie(S3)) or (EstSerie(S5))) then TS2.TabVisible:=False ;*)
if ((EstSerie(S3)) or (EstSerie(S5))) then EnableCeQuiFaut ;
if EstSerie(S3) then GQte.Visible:=False ;

END ;

procedure TFSaisComp.RemplitLesZones ;
var
 C : TComponent ;
 VV : Variant ;
 i : integer ;
 BlocNote : HTStringList ;
begin
  if RC.TOBCompl <> nil then
   begin
     {JP 24/10/05 : En saisie des règlement, avec contrepartie en pied de pièce, on arrive dans ce cas
                    avec un conversion de variant incorrect
     CUTOFFDEB.Text := DateToStr(RC.TOBCompl.GetValue('EC_CUTOFFDEB') ) ;
     CUTOFFFIN.Text := DateToStr(RC.TOBCompl.GetValue('EC_CUTOFFFIN') ) ;}
     if RC.TOBCompl.GetString('EC_CUTOFFDEB') = '' then CUTOFFDEB.Text := DateToStr(iDate1900)
                                                   else CUTOFFDEB.Text := DateToStr(RC.TOBCompl.GetValue('EC_CUTOFFDEB'));
     if RC.TOBCompl.GetString('EC_CUTOFFFIN') = '' then CUTOFFFIN.Text := DateToStr(iDate1900)
                                                   else CUTOFFFIN.Text := DateToStr(RC.TOBCompl.GetValue('EC_CUTOFFFIN'));
     if (CUTOFFDEB.Text = DateToStr(iDate1900)) and (CUTOFFFIN.Text = DateToStr(iDate1900)) then
      RC.AvecCalcul := true ;
   end ;
   for i:=0 to ComponentCount-1 do
       BEGIN
       C:=Components[i] ;
       if ((Copy(C.Name,1,1)='_') and (C.Tag<>-1)) then
          BEGIN
          VV:=OB.GetValue(Pf+C.Name) ;
          if C.ClassType=THNumEdit then THNumEdit(C).Value:=VarAsType(VV,varDouble) else
          if C.ClassType=THCpteEdit then THCpteEdit(C).Text:=VarAsType(VV,varString) else
          if C.ClassType=TEdit then TEdit(C).Text:=VarAsType(VV,varString) else
          if C.ClassType=TCheckBox then TCheckBox(C).Checked:=(VarAsType(VV,varString)='X') else
          if C.ClassType=TMaskEdit then TMaskEdit(C).Text:=DateToStr(VarAsType(VV,VarDate)) else
          if C.ClassType=THCritMaskEdit then THCritMaskEdit(C).Text:=DateToStr(VarAsType(VV,VarDate)) else   {FP FQ15916}
          // 13784
          if C is THRichEdit then
             begin
             BlocNote := HTStringList.Create;
             BlocNote.Text := VV;
             StringsToRich(THRichEdit(C),BlocNote);
             FreeAndNil(BlocNote);
             end else
          if C.ClassType=THValComboBox then THValComboBox(C).Value:=VarAsType(VV,VarString) ;
          if ((TControl(C).Parent=TS2) and (Action=taConsult)) then TControl(C).Enabled:=False ;
          ValeurDefModif(C) ;
          END ;
       if ((C.ClassType=TGroupBox) and (Action=taConsult)) then TControl(C).Enabled:=False ;
       END ;

end ;

procedure TFSaisComp.FormShow(Sender: TObject);
Var
    sTable : string ;
begin
case QuelEcr of
   EcrGen : sTable := 'E' ;
   EcrAna : sTable := 'Y' ;
   EcrBud : sTable := 'BE' ;
   end ;
FBoCutOff := Pos('CUT',RC.StLibre) > 0 ;
{$IFDEF CCS3}
_CONSO.Visible:=False ; H_CONSO.Visible:=False ;
{$ELSE}
if not FBoCutOff then PersoChampsLibres(sTable, ts2,true) ;
{$ENDIF}
ChangeLesMask ; InitTitre ; InitZones ;
RemplitLesZones ;
EtudieQtes ;
EtudieComporte ;
Case Action of
   taCreat   : Caption:=HC.Mess[1] ;
   taModif   : Caption:=HC.Mess[2] ;
   taConsult : Caption:=HC.Mess[3] ;
   END ;
if ((EstSerie(S7)) and (QuelEcr=EcrAna) and (LibreAnaObli)) then
   BEGIN
   if TS2.TabVisible then BEGIN Pages.ActivePage:=TS2 ; if _TABLE0.CanFocus then _TABLE0.SetFocus ; END ;
   END ;
if FBoCutOff and CUTOFFDEB.CanFocus then
 CUTOFFDEB.SetFocus
  else
   if FBoCutOff and CUTOFFFIN.CanFocus then
    CUTOFFFIN.SetFocus
     else
      if ((FocusBN) and (_BLOCNOTE.CanFocus)) then _BLOCNOTE.SetFocus ;
GeneModif:=False ;
end;

Function TFSaisComp.ConvertNom ( Nam : String ) : String ;
Var LeNom : String ;
BEGIN
LeNom:='' ;
if RC.Origine=0 then
   BEGIN
   if ((Nam='_REFINTERNE') or (Nam='_REFEXTERNE') or (Nam='_LIBELLE') or (Nam='_DATEREFEXTERNE') or
       (Nam='_REFLIBRE') or (Nam='_AFFAIRE')) then LeNom:='SC'+Nam
       else
   if ((Nam>='_LIBRETEXTE0') and (Nam<='_LIBRETEXTE9')) then LeNom:='ET'+Inttostr(Ord(Nam[12])-47)
       else
   if ((Nam>='_TABLE0') and (Nam<='_TABLE3')) then LeNom:='ET'+Inttostr(Ord(Nam[7])-37)
       else
   if ((Nam>='_LIBREMONTANT0') and (Nam<='_LIBREMONTANT3')) then LeNom:='ET'+Inttostr(Ord(Nam[14])-33)
       else
   if Nam='_LIBREBOOL0' then LeNom:='ET19'
       else
   if Nam='_LIBREBOOL1' then LeNom:='ET20'
       else
   if Nam='_LIBREDATE' then LeNom:='ET21' ;
   END else
   BEGIN
   if Nam='_REFEXTERNE' then LeNom:='LE3' else
   if Nam='_REFEXTERNE' then LeNom:='LE4' else
   if Nam='_AFFAIRE'    then LeNom:='LE5' else
   if Nam='_REFLIBRE'   then LeNom:='LE6' else
   if Nam='_QTE1'       then LeNom:='LE7' else
   if Nam='_QTE2'       then LeNom:='LE8' else
   if Nam='_BLOCNOTE'   then LeNom:='LE9' else
   if ((Nam>='_LIBRETEXTE0') and (Nam<='_LIBRETEXTE9')) then LeNom:='ZLL'+Inttostr(Ord(Nam[12])-47)
       else
   if ((Nam>='_TABLE0') and (Nam<='_TABLE3')) then LeNom:='ZLL'+Inttostr(Ord(Nam[7])-37)
       else
   if ((Nam>='_LIBREMONTANT0') and (Nam<='_LIBREMONTANT3')) then LeNom:='ZLL'+Inttostr(Ord(Nam[14])-33)
       else
   if Nam='_LIBREBOOL0' then LeNom:='ZLL19'
       else
   if Nam='_LIBREBOOL1' then LeNom:='ZLL20'
       else
   if Nam='_LIBREDATE' then LeNom:='ZLL21' ;
   END ;
Result:=LeNom ;
END ;

Function TFSaisComp.ZoneObli ( Nam : String ) : boolean ;
Var StOrig,Champ,Oblg,Modif,ValDef,TexteLibre : String ;
BEGIN
Result:=False ;
if RC.Origine<0 then Exit ;
if RC.Origine=0 then StOrig:='SC_LIBREENTETE' else StOrig:='SC_RADICAL'+Inttostr(RC.Origine) ;
Champ:=ConvertNom(Nam) ; if Champ='' then Exit ;
ValeurAttributComp(RC.MemoComp,StOrig,Champ,Oblg,Modif,Valdef,TexteLibre) ;
Result:=((Oblg='X') and (Modif='X')) ;
END ;

Function TFSaisComp.ZoneVide ( C : TComponent ) : boolean ;
BEGIN
Result:=True ;
if C.ClassType=THNumEdit then if THNumEdit(C).Value=0 then Exit ;
if C.ClassType=THCpteEdit then if THCpteEdit(C).Text='' then Exit ;
if C.ClassType=TEdit then if TEdit(C).Text='' then Exit ;
if C.ClassType=TMaskEdit then if TMaskEdit(C).Text=stDate1900 then Exit ;
if C.ClassType=THCritMaskEdit then if THCritMaskEdit(C).Text=stDate1900 then Exit ;   {FP FQ15916}
if C is TRichEdit then if TRichEdit(C).Lines.Count<=0 then Exit ;
Result:=False ;
END ;

Procedure TFSaisComp.ValeurDefModif ( C : TComponent ) ;
Var StOrig,Champ,Oblg,Modif,ValDef,TexteLibre : String ;
    DD : TDateTime ; 
BEGIN
if Action<>taCreat then Exit ;
if Not RC.Attributs then Exit ;
if ((Copy(C.Name,1,1)<>'_') or (C.Tag=-1)) then Exit ;
if RC.Origine<0 then Exit ;
if RC.Origine=0 then StOrig:='SC_LIBREENTETE' else StOrig:='SC_RADICAL'+Inttostr(RC.Origine) ;
Champ:=ConvertNom(C.Name) ; if Champ='' then Exit ;
ValeurAttributComp(RC.MemoComp,StOrig,Champ,Oblg,Modif,Valdef,TexteLibre) ;
if Modif<>'X' then TWinControl(C).Enabled:=False ;
if ValDef<>'' then
   BEGIN
   if C.ClassType=THNumEdit then THNumEdit(C).Value:=Valeur(ValDef) else
   if C.ClassType=THCpteEdit then THCpteEdit(C).Text:=ValDef else
   if C.ClassType=TEdit then TEdit(C).Text:=ValDef else
   if C.ClassType=TCheckBox then TCheckBox(C).Checked:=(ValDef='X') ;
   if (C.ClassType=TMaskEdit) or (C.ClassType=THCritMaskEdit) then  {FP FQ15916}
      BEGIN
      DD:=0 ;
      if ValDef='DC' then DD:=RC.DateC else
      if ValDef='DS' then DD:=Date else
      if ValDef='DE' then DD:=V_PGI.DateEntree else
      if ValDef='DD' then DD:=VH^.Encours.Deb else
      if ValDef='DF' then DD:=VH^.Encours.Fin ;
      {b FP FQ15916}
      if (DD>0) and (C.ClassType=TMaskEdit) then TMaskEdit(C).Text:=DateToStr(DD) ;
      if (DD>0) and (C.ClassType=THCritMaskEdit) then THCritMaskEdit(C).Text:=DateToStr(DD) ;
      {e FP FQ15916}
      END ;
   END ;
END ;

Function TFSaisComp.VerifObligatoire : boolean ;
Var i : integer ;
    C : TComponent ;
    CC : TControl ;
BEGIN
  {JP 18/10/05 : FQ 16899 : En CutOff avec écran réduit, on ne teste pas les champs obligatoires}
  if FBoCutOff then begin
    Result:=True ;
    Exit;
  end;

Result:=False ;
for i:=0 to ComponentCount-1 do
    BEGIN
    C:=Components[i] ;
    if ((Copy(C.Name,1,1)='_') and (C.Tag<>-1)) then
       BEGIN
//       if Not TWinControl(C).CanFocus then Continue ;
       CC:=TControl(C) ;
       if ZoneObli(C.Name) And (C<>NIL) And (CC.Enabled) then if ZoneVide(C) then
          BEGIN
          If C.Tag<100 Then Pages.ActivePage:=TS1 Else Pages.ActivePage:=TS2 ;
          TWinControl(C).SetFocus ;
          HM.Execute(0,'','') ; Exit ;
          END ;
       END ;
    END ;
Result:=True ;
END ;

////////////////////////////////////////////////////////////////////////////////
procedure TFSaisComp.ValideLesComps ;
var C : TComponent ;
    VV : Variant ;
    i : integer ;
    BlocNote : HTStringList ;
begin
  if Action=taConsult then Exit ;

  NextControl(Self); // FQ 18182 : SBO 07/06/2006

  if RC.TOBCompl <> nil then
   begin
    //CUTOFFDEBExit(nil) ;
    if ( StrToDate(CUTOFFDEB.Text) > StrToDate(CUTOFFFIN.Text) ) then
     begin
      PGIInfo('La date de début de charge périodique ne peut être supérieure à la date de fin !') ;
      CUTOFFFIN.Text := DateToStr(RC.TOBCompl.GetValue('EC_CUTOFFFIN') ) ;
      if CUTOFFFIN.CanFocus then CUTOFFFIN.SetFocus ;
      exit ;
     end ;
    if (StrToDate(CUTOFFDEB.Text)<>RC.TOBCompl.GetValue('EC_CUTOFFDEB')) or (StrToDate(CUTOFFFIN.Text)<>RC.TOBCompl.GetValue('EC_CUTOFFFIN')) then
     RC.TOBCompl.PutValue('EC_CUTOFFDATECALC', iDate1900 ) ;
    RC.TOBCompl.PutValue('EC_CUTOFFDEB', StrToDate(CUTOFFDEB.Text) ) ;
    RC.TOBCompl.PutValue('EC_CUTOFFFIN', StrToDate(CUTOFFFIN.Text) ) ;
   end ;
  
  if ((RC.Attributs) and (RC.MemoComp<>Nil)) then
  begin
    if Not VerifObligatoire then Exit ;
  end;

  for i:=0 to ComponentCount-1 do
  begin
    C:=Components[i] ; VV:=#0 ;
    if ((Copy(C.Name,1,1)='_') and (C.Tag<>-1)) then
    begin
      if C.ClassType = THNumEdit then
         VV := THNumEdit(C).Value
      else
      if C.ClassType = THCpteEdit then
        VV := THCpteEdit(C).Text
      else
      if C.ClassType = TEdit then
        VV := TEdit(C).Text
      else
      if C.ClassType = TCheckBox then
      begin
        if TCheckBox(C).Checked then
          VV := 'X'
        else
          VV := '-' ;
      end
      else
      if C.ClassType = TMaskEdit then
        VV := StrToDate(TMaskEdit(C).Text)
      {b FP FQ15916}
      else
      if C.ClassType = THCritMaskEdit then
        VV := StrToDate(THCritMaskEdit(C).Text)
      {e FP FQ15916}
      else
      if C is TRichEdit then  // 13784
      begin
        BlocNote := HTStringList.Create ;
        RichToStrings(THRichEdit(C),BlocNote) ;
        if (THRichEdit(C).Text = '') then
          OB.PutValue(Pf+C.Name,'')      // FQ 10419
        else
          OB.PutValue(Pf+C.Name,BlocNote.Text);
        FreeAndNil(BlocNote) ;
      end
      else
      if C.ClassType = THValComboBox then
        VV := THValComboBox(C).Value ;

      if Not (C is TRichEdit) then OB.PutValue(Pf+C.Name,VV) ;
    end ;
  end ;
  //if RC.TOBCompl <> nil then
  // OkV := ( CUTOFFDEB.Text <> DateToStr(iDate1900) ) and ( CUTOFFFIN.Text <> DateToStr(iDate1900) )
  //  else
  OkV := True ;
  GeneModif := False ;
  Close ;
end;
////////////////////////////////////////////////////////////////////////////////

procedure TFSaisComp.BValideClick(Sender: TObject);
begin
  ValideLesComps ;
end;

procedure TFSaisComp._QTE1Exit(Sender: TObject);
Var X : Double ;
begin
if QuelEcr in [EcrGen,EcrBud] then Exit ;
X:=Valeur(TotQ1.Text) ; if X<=0 then Exit ;
_POURCENTQTE1.Text:=StrfMontant(100.0*Valeur(_QTE1.Text)/X,15,ADecimP,'',True) ;
end;

procedure TFSaisComp._QTE2Exit(Sender: TObject);
Var X : Double ;
begin
if QuelEcr in [EcrGen,EcrBud] then Exit ;
X:=Valeur(TotQ2.Text) ; if X<=0 then Exit ;
_POURCENTQTE2.Text:=StrfMontant(100.0*Valeur(_QTE2.Text)/X,15,ADecimP,'',True) ;
end;

procedure TFSaisComp._POURCENTQTE1Exit(Sender: TObject);
Var X : Double ;
begin
if QuelEcr in [EcrGen,EcrBud] then Exit ;
X:=Valeur(TotQ1.Text) ; if X<=0 then Exit ;
_QTE1.Text:=StrfMontant(X*Valeur(_POURCENTQTE1.Text)/100.0,15,V_PGI.OkDecQ,'',True) ;
end;

procedure TFSaisComp._POURCENTQTE2Exit(Sender: TObject);
Var X : double ;
begin
if QuelEcr in [EcrGen,EcrBud] then Exit ;
X:=Valeur(TotQ2.Text) ; if X<=0 then Exit ;
_QTE2.Text:=StrfMontant(X*Valeur(_POURCENTQTE2.Text)/100.0,15,V_PGI.OkDecQ,'',True) ;
end;

procedure TFSaisComp._BLOCNOTEChange(Sender: TObject);
begin
ModifBN:=True ;
GeneModif:=True ;
end;

Function ComporteLigne ( Scenario : TOBM ; General : String17 ; Var StComporte,StLibre : String ) : integer ;
Var i : integer ;
    Rad : String ;
BEGIN
StComporte:='' ; StLibre:='' ; Result:=0 ;
if Scenario=Nil then Exit ; if General='' then Exit ;
for i:=1 to 10 do
    BEGIN
    Rad:=Scenario.GetValue('SC_RADICAL'+IntToStr(i)) ;
    if ((Rad<>'') and (Copy(General,1,Length(Rad))=Rad)) or (Rad='*') then
       BEGIN
       StComporte:=Scenario.GetValue('SC_COMPLEMENTS'+IntToStr(i)) ;
       StLibre:=Scenario.GetValue('SC_COMPLIBRE'+IntToStr(i)) ;
       Result:=i ;
       Break ;
       END ;
    END ;
END ;

procedure TFSaisComp.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var Vide : boolean ;
begin
Vide:=(Shift=[]) ;
Case Key of
   VK_F10    : BEGIN Key:=0 ; if Vide then ValideLesComps ; END ;
   VK_ESCAPE : BEGIN Key:=0 ; if Vide then BAbandonClick(Nil) ; END ;
   END ;
end;

procedure TFSaisComp.BAbandonClick(Sender: TObject);
begin
  {JP 27/10/05 : Comme les SpeedButtons ne prennent pas le focus, on force la sortie
                 du control actif}
  NextControl(Self);
  Close ;
end;

procedure TFSaisComp.FormCreate(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ; OkV:=False ;
//WMinX:=Width ; WMinY:=Height ;
GeneModif:=False ;
end;

{
procedure TFSaisComp.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;}

procedure TFSaisComp._REFEXTERNEChange(Sender: TObject);
begin
GeneModif:=True ;
end;

procedure TFSaisComp.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
Var ii : integer ;
begin
if Action=taConsult then Exit ;
if ((RC.Attributs) and (RC.MemoComp<>Nil)) then
   if Not VerifObligatoire then BEGIN CanClose:=False ; Exit ; END ;
if Not GeneModif then Exit ;
ii:=HC.Execute(5,'','') ;
Case ii of
   mrYes : ValideLesComps ;
   mrNo  : GeneModif:=False ;
   mrCancel : CanClose:=False ;
   END ;
end;

procedure TFSaisComp.BAideClick(Sender: TObject);
begin
  {JP 27/10/05 : Comme les SpeedButtons ne prennent pas le focus, on force la sortie
                 du control actif}
  NextControl(Self);
  CallHelpTopic(Self) ;
end;

procedure TFSaisComp.VerifLibre(Sender: TObject);
Var HH : THCpteEdit ;
		Test : Boolean ;
begin
  HH:=THCpteEdit(Sender) ;
  if HH.Text='' then Exit ;
  Test := GChercheCompte(HH,Nil);
  if Not Test
	  then HH.Text:='' ;
end;

procedure TFSaisComp.DBCPerso(Sender: TObject);
Var StZ,StOrig,Champ,Oblg,Modif,ValDef,TexteLibre : String ;
begin
if Action<>taCreat then Exit ;
if Not RC.Attributs then Exit ;
if ((Copy(TControl(Sender).Name,1,1)<>'_') or (TControl(Sender).Tag=-1)) then Exit ;
if RC.Origine<0 then Exit ;
if RC.Origine=0 then StOrig:='SC_LIBREENTETE' else StOrig:='SC_RADICAL'+Inttostr(RC.Origine) ;
Champ:=ConvertNom(TControl(Sender).Name) ; if Champ='' then Exit ;
ValeurAttributComp(RC.MemoComp,StOrig,Champ,Oblg,Modif,Valdef,TexteLibre) ;
if TexteLibre='' then Exit ;
StZ:=Trim(uppercase(TexteLibre)) ;
if StZ='TZGENERAL'  then HCptPerso.ZoomTable:=tzGeneral  else
if StZ='TZTIERS'    then HCptPerso.ZoomTable:=tzTiers    else
if StZ='TZJOURNAL'  then HCptPerso.ZoomTable:=tzJournal  else
if StZ='TZSECTION'  then HCptPerso.ZoomTable:=tzSection  else
if StZ='TZSECTION2' then HCptPerso.ZoomTable:=tzSection2 else
if StZ='TZSECTION3' then HCptPerso.ZoomTable:=tzSection3 else
if StZ='TZSECTION4' then HCptPerso.ZoomTable:=tzSection4 else
if StZ='TZSECTION5' then HCptPerso.ZoomTable:=tzSection5 else
if StZ='TZBUDGEN'   then HCptPerso.ZoomTable:=tzBudGen   else
if StZ='TZBUDSEC1'  then HCptPerso.ZoomTable:=tzBudSec1  else
if StZ='TZBUDSEC2'  then HCptPerso.ZoomTable:=tzBudSec2  else
if StZ='TZBUDSEC3'  then HCptPerso.ZoomTable:=tzBudSec3  else
if StZ='TZBUDSEC4'  then HCptPerso.ZoomTable:=tzBudSec4  else
if StZ='TZBUDSEC5'  then HCptPerso.ZoomTable:=tzBudSec5  else Exit ;
HCptPerso.Text:=TEdit(Sender).Text ;
if GChercheCompte(HCptPerso,Nil) then TEdit(Sender).Text:=HCptPerso.Text ;
end;

procedure TFSaisComp._AFFAIREKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_F5 then GereAffaire;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 21/02/2002
Modifié le ... :   /  /
Description .. : Modif pour le nouveau type OBM :
Suite ........ : le tableau de T[] n'existe plus
Mots clefs ... :
*****************************************************************}
procedure TFSaisComp.GereAffaire;
var stAffaire, stAuxi : string;
//  i : integer;
begin
  // Appel du multicritères affaire uniquement si saisie sur GI (même base (MEM) ou base commune (COM))
  if ((GetParamSoc('SO_CPSAISIEAFFAIRE')='MEM') or (GetParamSoc('SO_CPSAISIEAFFAIRE')='COM')) then
  begin
   { i := -1;
    case OB.Ident of
      EcrGen : i := TrouveIndice(VH^.DescriEcr,'E_AUXILIAIRE',True);
      EcrBud : i := TrouveIndice(VH^.DescriEcr,'E_AUXILIAIRE',True);
      EcrAna : i := TrouveIndice(VH^.DescriAna,'Y_AUXILIAIRE',True);
    end;
    if ((i>=0) and (i<=High(OB.T))) then stAuxi := OB.T[i].V;
    if stAuxi = '' then exit;  }
    case OB.Ident of
      EcrGen, EcrBud : StAuxi:=OB.GetValue('E_AUXILIAIRE');
      EcrAna         : StAuxi:=OB.GetValue('Y_AUXILIAIRE');
    end;
    if stAuxi = '' then exit;
    stAffaire := AGLLanceFiche ('AFF','AFFAIRE_MUL_COMPT','AFF_STATUTAFFAIRE=AFF','','NOCHANGESTATUT;AUXI='+stAuxi);
    if stAffaire <> '' then _AFFAIRE.Text := stAffaire;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 10/10/2005
Modifié le ... : 19/10/2005
Description .. : - LG - 10/10/2005 - on ne calcul la date de fin que pour le 
Suite ........ : termne echue
Suite ........ : - LG - 19/10/2005 - FB 16916 - J'ai mis en périodicité 
Suite ........ : Année	Début : 01/01/2005 ==> Fin : 31/12/2005
Suite ........ : Lorsque je retourne dans les informations complémentaires 
Suite ........ : si j'utilise une autre zone (en faisant Tabulation par 
Suite ........ : exemple), la date de fin retourne au 31/03/2005 
Suite ........ : systématiquement
Mots clefs ... : 
*****************************************************************}
procedure TFSaisComp.CUTOFFFINExit(Sender: TObject);
begin
 if RC.AvecCalcul then
  begin
   RC.TOBCompl.PutValue('EC_DATECOMPTABLE', StrToDate(CUTOFFFIN.Text) ) ;
   CalculDateCutOff ;
  end ;
 if _BLOCNOTE.CanFocus then _BLOCNOTE.SetFocus ;
end;

procedure TFSaisComp.BDetailClick(Sender: TObject);
begin
  {JP 27/10/05 : Comme les SpeedButtons ne prennent pas le focus, on force la sortie
                 du control actif}
  NextControl(Self);
  BDetail.Visible  := not BDetail.Visible ;
  BReduire.Visible := not BDetail.Visible ;
  PRef.Visible     := BDetail.Visible ;
  PInfo.Visible    := BDetail.Visible ;
  PCutOff.Visible  := FBoCutOff ;
  TS1.TabVisible   := BDetail.Visible ;
  TS2.TabVisible   := BDetail.Visible ;

  if BDetail.Visible then
  begin
   Self.Height           := 475 ;
   Pages.ActivePageIndex := 0 ;
  end
  else
  begin
    Self.Height           := 250 ;
    Pages.ActivePageIndex := 0 ;
  end ;
end;

procedure TFSaisComp.BReduireClick(Sender: TObject);
begin
 BDetailClick(Sender) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 25/07/2005
Modifié le ... : 10/05/2006
Description .. : LG - 25/07/2005 - FB 15874 - voir fiche
Suite ........ : - LG  - 13/10/2005 -  FB 16669 - correction de l'interaction
Suite ........ : en la periode et le declenchemnt du calcul de la periode
Suite ........ : - LG - 19/10/2005 - FB 16916 - J'ai mis en périodicité
Suite ........ : Année	Début : 01/01/2005 ==> Fin : 31/12/2005
Suite ........ : Lorsque je retourne dans les informations complémentaires
Suite ........ : si j'utilise une autre zone (en faisant Tabulation par
Suite ........ : exemple), la date de fin retourne au 31/03/2005
Suite ........ : systématiquement
Suite ........ : - LG - 10/05/2006 - FB 17998 - on prends la date saisie 
Suite ........ : pour effectuer la calcul de cutoff
Mots clefs ... : 
*****************************************************************}
procedure TFSaisComp.CUTOFFDEBExit(Sender: TObject);
begin
 if CUTOFFDEB.Text = DateToStr(iDate1900) then
  CUTOFFFIN.Text := DateToStr(iDate1900) ;
 if RC.AvecCalcul then
  begin
   RC.TOBCompl.PutValue('EC_DATECOMPTABLE', StrToDate(CUTOFFDEB.Text) ) ;
   CalculDateCutOff ;
  end ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 25/07/2005
Modifié le ... : 20/01/2006
Description .. : LG - 25/07/2005 - FB 15874 - voir fiche
Suite ........ : - LG  - 13/10/2005 -  FB 16669 - correction de l'interaction
Suite ........ : en la periode et le declenchemnt du calcul de la periode
Suite ........ : - LG - 20/01/2005 - FB 16968 - on n'affecte plus la date
Suite ........ : comptable avec la date de cutoff calculer
Mots clefs ... :
*****************************************************************}
procedure TFSaisComp.CalculDateCutOff ;
begin

 if ( RC.TOBCompl = nil ) then exit ;

 if CUTOFFPERIODE.Value = 'CN' then exit ;

 CCalculDateCutOff( RC.TOBCompl , CUTOFFPERIODE.Value , RC.CutOffEchue ) ;
 if RC.CutOffEchue = '-' then
  CUTOFFFIN.Text := DateToStr(RC.TOBCompl.GetValue('EC_CUTOFFFIN') )
   else
    CUTOFFDEB.Text := DateToStr(RC.TOBCompl.GetValue('EC_CUTOFFDEB') ) ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 25/07/2005
Modifié le ... : 30/08/2005
Description .. : LG - 25/07/2005 - FB 15874 - voir fiche
Mots clefs ... : 
*****************************************************************}
procedure TFSaisComp.CUTOFFPERIODEChange(Sender: TObject);
begin
 CalculDateCutOff ;
 CUTOFFFIN.Enabled := ( CUTOFFPERIODE.Text = 'Aucun' ) or ( RC.CutOffEchue = 'X' ) ;
 CUTOFFDEB.Enabled := ( CUTOFFPERIODE.Text = 'Aucun' ) or  ( RC.CutOffEchue = '-' ) ;
end;

{JP 08/08/05 : FQ 14919 : Déplacement de l'appel à GereAffaire qui était dans DBCPerso
               et qui pouvait être appelé sur un double clik
{---------------------------------------------------------------------------------------}
procedure TFSaisComp._AFFAIREDblClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {Appel à la gestion d'affaire}
  GereAffaire;
  {Traitement générique sur les différentes zones de la saisie complémentaire}
  DBCPerso(Sender);
end;

end.
