unit AssistGenereCAB;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HSysMenu, hmsgbox, StdCtrls, HTB97, ComCtrls, ExtCtrls, Hctrls,
  UTob,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HStatus,UtilGC,Buttons,HEnt1,ParamSoc, HPanel ;

  procedure Assist_GenereCAB(TobA : TOB) ;

type
  TFAssistGenereCAB = class(TFAssist)
    GBcab: TGroupBox;
    Panel1: TPanel;
    rbMajBlanc: TRadioButton;
    rbMajTous: TRadioButton;
    rbEfface: TRadioButton;
    HLabel1: THLabel;
    procedure FormShow(Sender: TObject) ;
    procedure bFinClick(Sender: TObject) ;
    procedure MajCAB(MajBlanc,MajTous,MajEfface : Boolean) ;
    procedure CreateWindowProgress(WindowCaption,LabelCaption,InfoInit,BtnCaption:String) ;
    procedure bAnnulerClick(Sender: TObject);
  private
    { Déclarations privées }
    TobArticle: TOB ;
  public
    { Déclarations publiques }
    ArreterClick : boolean;
    // Contrôles fenêtre de progression
    Animate1: TAnimate;
    PB1: TProgressBar;
    Label1: TLabel;
    Info: TLabel;
    bAnnuler: TBitbtn;
    TWSUPPR: TToolWindow97;
  end;

const
	// libellés des messages
	TexteMessage: array[1..7] of string 	= (
          {1}         'Confirmation'
          {2}        ,'Les codes à barres existants des articles sélectionnés vont être modifiés !'
          {3}        ,'Les codes à barres existants des articles sélectionnés vont être effacés !'
          {4}        ,'Génération des codes à barres'
          {5}        ,'Article'
          {6}        ,'Préparation du traitement ...'
          {7}        ,'&Arrêter'
                     );

var
  FAssistGenereCAB: TFAssistGenereCAB;

implementation

{$R *.DFM}

procedure Assist_GenereCAB(TobA:TOB) ;
begin
FAssistGenereCAB:=TFAssistGenereCAB.Create(Application) ;
FAssistGenereCAB.TOBArticle:=TobA ;
Try
   FAssistGenereCAB.ShowModal ;
Finally
   FAssistGenereCAB.Free ;
end ;
end ;

procedure TFAssistGenereCAB.FormShow(Sender: TObject) ;
begin
inherited;
bPrecedent.Visible:=False ;
bSuivant.Visible:=False ;
lAide.Caption:='' ;

// Création d'une fenêtre de progression
CreateWindowProgress(TexteMessage[4],TexteMessage[5],TexteMessage[6],TexteMessage[7]) ;
end ;

procedure TFAssistGenereCAB.CreateWindowProgress(WindowCaption,LabelCaption,InfoInit,
                                                 BtnCaption:String) ;
begin
TWSUPPR:=TToolWindow97.Create(Self) ;
TWSUPPR.Parent:=Self ;
TWSUPPR.Caption:=WindowCaption ;
TWSUPPR.Visible:=False ;
TWSUPPR.Width:=320 ;
TWSUPPR.Height:=170 ;
TWSUPPR.Top:=Self.Top+Trunc((Self.Height-TWSUPPR.Height)/2) ;
TWSUPPR.Left:=Self.Left+Trunc((Self.Width-TWSUPPR.Width)/2) ;
Animate1:=TAnimate.Create(TWSUPPR) ;
Animate1.Parent:=TWSUPPR ;
Animate1.ParentWindow:=TWSUPPR.ParentWindow ;
Animate1.CommonAVI:=aviCopyFiles ;
Animate1.Align:=alTop ;
Animate1.Center:=True ;
Animate1.Active:=True ;
Label1:=TLabel.Create(TWSUPPR) ;
Label1.Parent:=TWSUPPR ;
Label1.Caption:=LabelCaption ;
Label1.Top:=Animate1.Top+Animate1.Height+10 ;
Label1.Left:=5;
Info:=TLabel.Create(TWSUPPR) ;
Info.Parent:=TWSUPPR ;
Info.Caption:=InfoInit ;
Info.Top:=Label1.Top;
Info.Left:=Label1.Left+Label1.Width+5 ;
PB1:=TProgressBar.Create(TWSUPPR) ;
PB1.ParentWindow:=TWSUPPR.ParentWindow ;
PB1.Parent:=TWSUPPR ;
PB1.Width:=TWSUPPR.ClientWidth-10 ;
PB1.Min:=0 ;
PB1.Max:=100 ;
PB1.Left:=5 ;
PB1.Top:=Label1.Top+Label1.Height+10 ;
bAnnuler:=TBitBtn.Create(TWSUPPR) ;
bAnnuler.Parent:=TWSUPPR ;
bAnnuler.OnClick:=bAnnulerClick ;
bAnnuler.Caption:=BtnCaption ;
bAnnuler.Top:=PB1.Top+PB1.Height+10 ;
bAnnuler.Left:=Trunc((TWSUPPR.ClientWidth-bAnnuler.Width)/2) ;
end ;

procedure TFAssistGenereCAB.bAnnulerClick(Sender: TObject);
begin
// DCA - FQ MODE 10021 - Btn Annuler
inherited;
ArreterClick:=True ;
end ;

// Traitement de maj des Codes à Barres
procedure TFAssistGenereCAB.MajCAB(MajBlanc,MajTous,MajEfface : Boolean) ;
var TobMulArt,TobArt,TobDim : TOB ;
    iMul,iArt : integer ;
    bMajCAB,bMajAuto,bFirst,bCtrl : boolean ;
    QQ : TQuery ;
    SQL,FournPrinc,CodeCAB,TypeCAB,Titre : string ;
begin
TobMulArt:=TOBArticle ;

// Gestion fenêtre de progression
if rbMajBlanc.Checked then Titre:=rbMajBlanc.Caption
else if rbMajTous.Checked then Titre:=rbMajTous.Caption
else if rbEfface.Checked then Titre:=rbEfface.Caption ;
TWSUPPR.Caption:=Titre ;
PB1.Position:=0 ;
if TobMulArt.Detail.Count>0 then PB1.Step:=Trunc(100/TobMulArt.Detail.Count) ;
TWSUPPR.Show ;

TobMulArt.Detail.Sort('GA_FOURNPRINC;GA_ARTICLE') ;
FournPrinc:='' ; TypeCAB:='' ; //CodeCAB:='' ; 
bMajAuto:=False ; bFirst:=True ;

BeginTrans ;
SQL:='select GA_ARTICLE,GA_FOURNPRINC,GA_CODEBARRE,GA_QUALIFCODEBARRE,GA_STATUTART ' +
     'from ARTICLE where GA_CODEARTICLE=:CODEARTICLE' ;
QQ:=PrepareSQL(SQL,True) ;
ArreterClick:=False ;
for iMul:=0 to TobMulArt.Detail.Count-1 do
    BEGIN
    With TobMulArt.Detail[iMul] do
        BEGIN
        Application.ProcessMessages ;
        if ArreterClick then Break ;
        Info.Caption:=GetValue('GA_CODEARTICLE') ;

        // Articles génériques, dimensionnés et tailles uniques
        //SQL:='select GA_ARTICLE,GA_FOURNPRINC,GA_CODEBARRE,GA_QUALIFCODEBARRE,GA_STATUTART ' +
        //     'from ARTICLE where GA_CODEARTICLE="'+GetValue('GA_CODEARTICLE')+'"' ;
        //QQ:=OpenSQL(SQL,True) ;
        QQ.ParamByName('CODEARTICLE').AsString:=GetValue('GA_CODEARTICLE') ;
        QQ.Open ;
        if QQ<>nil then
            BEGIN
            TobArt:=TOB.CreateDB('ARTICLE',nil,-1,QQ) ;
            TobArt.LoadDetailDB('ARTICLE','','',QQ,False) ;
            if TobArt<>nil then
                BEGIN
                for iArt:=0 to TobArt.Detail.Count-1 do
                    BEGIN
                    TobDim:=TobArt.Detail[iArt] ;
                    if ((MajBlanc) and (   ((TobDim.GetValue('GA_CODEBARRE')='') and (TobDim.GetValue('GA_STATUTART')<>'GEN'))
                                        or ((TobDim.GetValue('GA_QUALIFCODEBARRE')='') and (TobDim.GetValue('GA_STATUTART')='GEN')) ))
                        or (MajTous) then
                        BEGIN
                        if (TobDim.GetValue('GA_FOURNPRINC')<>FournPrinc) or (bFirst) then
                            BEGIN
                            //Géré dans AttribNewCodeCAB !! if (CodeCAB<>'') then SauvChronoCAB(FournPrinc) ;
                            FournPrinc:=TobDim.GetValue('GA_FOURNPRINC') ;
                            CodeCAB:='' ; bFirst:=False ;
                            bMajAuto:=RechParamCAB_Auto(FournPrinc,TypeCAB,True) ;
                            END ;
                        if bMajAuto then
                            BEGIN
                            bMajCAB:=True ;
                            if TobDim.GetValue('GA_STATUTART')<>'GEN' then
                                BEGIN
                                if TobDim.GetValue('GA_TYPEARTICLE')='PRE'
                                    then bCtrl:=GetParamSoc('SO_GCPRCABPRESTATION')
                                    else bCtrl:=GetParamSoc('SO_GCCABARTICLE') ;
                                CodeCAB := AttribNewCodeCAB ( bCtrl, FournPrinc ) ;
                                END ;
                            END
                            else bMajCAB:=False ;
                        END
                        else bMajCAB:=MajEfface ;
                    if bMajCAB then
                        BEGIN
                        if TobDim.GetValue('GA_STATUTART')<>'GEN' then TobDim.PutValue('GA_CODEBARRE',CodeCAB) ;
                        TobDim.PutValue('GA_QUALIFCODEBARRE',TypeCAB) ;
                        END ;
                    END ;
                TobArt.UpdateDB(False,True) ; // Effectue une requête uniquement si un champ de la tob a été modifié
                TobArt.Free ;
                END ;
            END ;
        QQ.Close ;
        END ;
    PB1.StepIt ;
    END ;
Ferme(QQ) ;
if Not ArreterClick then
    BEGIN
    //Géré dans AttribNewCodeCAB !! if (CodeCAB<>'') then SauvChronoCAB(FournPrinc) ;
    CommitTrans ;
    END
    else RollBack ;
TWSUPPR.Hide ;
end ;

procedure TFAssistGenereCAB.bFinClick(Sender: TObject);
begin
inherited ;
if rbMajTous.Checked then
    if PGIAsk(TexteMessage[2],TexteMessage[1])<>mrYes then exit ;
if rbEfface.Checked then
    if PGIAsk(TexteMessage[3],TexteMessage[1])<>mrYes then exit ;
MajCAB(rbMajBlanc.Checked,rbMajTous.Checked,rbEfface.Checked) ;
Close ;
end ;

end.
