{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 12/02/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : GCMULSUPPRART ()
Mots clefs ... : TOF;GCMULSUPPRART
*****************************************************************}
Unit UTofGcMulSupprArt ;

Interface

Uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls, Windows,
     ExtCtrls, Buttons, HPanel, HCtrls, HEnt1, HMsgBox, UTOB, UTOF, HTB97,
{$IFDEF EAGLCLIENT}
     MaineAGL, eMul,
{$ELSE}
     Mul, db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Fe_Main,
{$ENDIF}
{$IFDEF AFFAIRE}
     AfUtilArticle,
{$ENDIF}
     M3FP, UTofGcModifLot, UtilArticle, AGLInit,UtilGc;

Type
  TOF_GCMULSUPPRART = Class (TOF)
    Animate1: TAnimate;
    PB1: TProgressBar;
    Label1: TLabel;
    Art: TLabel;
    bAnnuler: TBitbtn;
    TWSUPPR: TToolWindow97;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure bAnnulerClick(Sender: TObject);
  public
    TobErreur : TOB;
    SelArt : TOB ;
    TobSupp : TOB;
    FirstPass : boolean;
    ArreterClick : boolean;
    procedure SupprimeArticle ;
    procedure TraiteSelection ;
//    function  VerificationArticle(Article,CodeArticle: String) : string ;
//    procedure SuppressionArticle (Article, CodeArticle : string);
  end ;

const
	// libellés des messages
	TexteMessage: array[1..2] of string 	= (
          {1}        'Information'
          {2}       ,'Les articles sélectionnés ont été supprimés'
                     );

Procedure AFLanceFiche_Mul_Suppr_Art;

Implementation

procedure TOF_GCMULSUPPRART.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_GCMULSUPPRART.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_GCMULSUPPRART.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_GCMULSUPPRART.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_GCMULSUPPRART.OnArgument (S : String ) ;
var
   TLab : THLabel;
{$ifdef AFFAIRE}
ComboTypeArticle:THValComboBox;
{$endif}
begin
Inherited ;

if ctxAffaire in V_PGI.PGIContexte  then begin
                  //article
  if ((Ecran.FindComponent ('GA_LIBREART1'))<>nil) then
      begin
      GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'GA_LIBREART', 10, '_');
      GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GA_VALLIBRE', 3, '_');
      GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GA_DATELIBRE', 3, '_');
      GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GA_CHARLIBRE', 3, '_');
      GCMAJChampLibre (TForm (Ecran), False, 'BOOL', 'GA_BOOLLIBRE', 3, '_');
{$ifdef AFFAIRE}
  //mcd 05/03/02
ComboTypeArticle:=THValComboBox(GetControl('GA_TYPEARTICLE'));
ComboTypeArticle.plus:=PlusTypeArticle;
{$endif}
      end;

  TLab:=THLabel (Ecran.FindComponent ('TGA_FAMILLENIV1'));
  if tlab <>NIL then SetCOntrolText('TGA_FAMILLENIV1', RechDom('GCLIBFAMILLE','LF1',false));
  TLab:=THLabel (Ecran.FindComponent ('TGA_FAMILLENIV2'));
  if tlab <>NIL then SetCOntrolText('TGA_FAMILLENIV2', RechDom('GCLIBFAMILLE','LF2',false));
  TLab:=THLabel (Ecran.FindComponent ('TGA_FAMILLENIV3'));
  if tlab <>NIL then SetCOntrolText('TGA_FAMILLENIV3', RechDom('GCLIBFAMILLE','LF3',false));
  end;
  
{$IFDEF BTP}
THValComboBox ( getcontrol('GA_TYPEARTICLE')).Plus :=
    'AND (CO_CODE="MAR" OR CO_CODE="POU" OR CO_CODE="PRE" OR CO_CODE="OUV" OR CO_CODE="ARP")';
THEdit(getcontrol('XX_WHERE')).Text := 'AND (GA_TYPEARTICLE="MAR" OR GA_TYPEARTICLE="POU" OR GA_TYPEARTICLE="PRE" OR GA_TYPEARTICLE="OUV" OR GA_TYPEARTICLE="ARP")';
{$ENDIF}

TWSUPPR := TToolWindow97.Create(TFMul(Ecran));
TWSUPPR.Parent := TFMul(Ecran);
TWSUPPR.Caption := 'Suppression en cours...';
TWSUPPR.Visible := False;
TWSUPPR.Width := 321;
TWSUPPR.Height := 170;
TWSUPPR.Top := TFMul(Ecran).FListe.Top + Trunc((TFMul(Ecran).FListe.Height - TWSUPPR.Height) / 2);
TWSUPPR.Left := TFMul(Ecran).FListe.Left + Trunc((TFMul(Ecran).FListe.Width - TWSUPPR.Width) / 2);
Animate1 := TAnimate.Create(TWSUPPR);
Animate1.Parent := TWSUPPR;
Animate1.ParentWindow := TWSUPPR.ParentWindow;
Animate1.CommonAVI := aviDeleteFile;
Animate1.Align := alTop;
Animate1.Center := True;
Animate1.Active := True;
Label1 := TLabel.Create(TWSUPPR);
Label1.Parent := TWSUPPR;
Label1.Caption := 'Suppression de l''article';
Label1.Top := Animate1.Top + Animate1.Height + 10;
Label1.Left := 5;
Art := TLabel.Create(TWSUPPR);
Art.Parent := TWSUPPR;
Art.Caption := '';
Art.Top := Label1.Top;
Art.Left := Label1.Left + Label1.Width + 5;
PB1 := TProgressBar.Create(TWSUPPR);
PB1.ParentWindow := TWSUPPR.ParentWindow;
PB1.Parent := TWSUPPR;
PB1.Width := TWSUPPR.ClientWidth - 10;
PB1.Min := 0;
PB1.Max := 100;
PB1.Left := 5;
PB1.Top := Label1.Top + Label1.Height + 10;
bAnnuler := TBitBtn.Create(TWSUPPR);
bAnnuler.Parent := TWSUPPR;
bAnnuler.OnClick:=bAnnulerClick;
bAnnuler.Caption := '&Arreter';
bAnnuler.Top := PB1.Top + PB1.Height + 10;
bAnnuler.Left := Trunc((TWSUPPR.ClientWidth - bAnnuler.Width) / 2);
//uniquement en line
//Setcontrolvisible('PCOMPLEMENT', False);
end ;

procedure TOF_GCMULSUPPRART.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_GCMULSUPPRART.bAnnulerClick(Sender: TObject);
begin
if not ArreterClick then ArreterClick := True;
end;

/////////////////// Appellé par le bouton Valider /////////////////
procedure AGLSupprimeArticle(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFMul) then TOTOF:=TFMul(F).LaTOF else exit;

if (TOTOF is TOF_GCMULSUPPRART) then TOF_GCMULSUPPRART(TOTOF).SupprimeArticle else exit;
end;

///////////////// Suppression par lot d'aricle //////////////////////
procedure TOF_GCMULSUPPRART.SupprimeArticle ;
var TobTemp : TOB;
    F : TFMul ;
    i,nbEnrgts : integer;
begin

F:=TFMul(Ecran);
if (F.FListe.NbSelected = 0) and ( not F.FListe.AllSelected) then
    begin
    PGIInfo('Aucun article n''a été sélectionné',Ecran.Caption);
    exit;
    end;
//  nbEnrgts:=0 ; result:=0;
if HShowMessage('0;Confirmation;Confirmez vous l''épuration des articles?;Q;YN;N;N;','','') <> mrYes then exit;
{$IFDEF EAGLCLIENT}
if F.bSelectAll.Down then
   if not F.Fetchlestous then
     begin
     F.bSelectAllClick(Nil);
     F.bSelectAll.Down := False;
     exit;
     end;
{$ENDIF}

TobErreur:= TOB.Create('les erreurs',NIL,-1);

//begintrans;
SelArt := TOB.Create('_REFERENCEMENT',Nil,-1);

//**** tout a été sélectionné avec le bouton tout sélectionner ****
if (F.FListe.AllSelected) then
begin
  F.Q.DisableControls;
  nbEnrgts:=F.Q.RecordCount; F.Q.First ;
  for i:=0 to nbEnrgts-1 do
    begin
      TobTemp := TOB.Create('', SelArt, -1);
      TobTemp.AddChampSup('GA_ARTICLE', False);
      TobTemp.PutValue('GA_ARTICLE',TFmul(Ecran).Q.FindField('GA_ARTICLE').AsString);
      TobTemp.AddChampSup('GA_CODEARTICLE', False);
      TobTemp.PutValue('GA_CODEARTICLE',TFmul(Ecran).Q.FindField('GA_CODEARTICLE').AsString);
      // DCA - FQ MODE 10815
      TobTemp.AddChampSupValeur('GA_STATUTART', TFmul(Ecran).Q.FindField('GA_STATUTART').AsString, False);
      F.Q.Next;
    end;
  F.Q.EnableControls;
end
else
begin
  nbEnrgts:=F.FListe.nbSelected ;
  for i:=0 to nbEnrgts-1 do
    begin
      F.FListe.GotoLeBOOKMARK(i);
{$IFDEF EAGLCLIENT}
      TFMul(Ecran).Q.TQ.Seek(F.FListe.Row-1) ;
{$ENDIF}
      TobTemp := TOB.Create('', SelArt, -1);
      TobTemp.AddChampSup('GA_ARTICLE', False);
      TobTemp.PutValue('GA_ARTICLE',TFmul(Ecran).Q.FindField('GA_ARTICLE').AsString);
      TobTemp.AddChampSup('GA_CODEARTICLE', False);
      TobTemp.PutValue('GA_CODEARTICLE',TFmul(Ecran).Q.FindField('GA_CODEARTICLE').AsString);
      // DCA - FQ MODE 10815
      TobTemp.AddChampSupValeur('GA_STATUTART', TFmul(Ecran).Q.FindField('GA_STATUTART').AsString, False);
    end;
end;

TraiteSelection;

//Ajout suite à synchro 24/09/2001 mais à tester !!!
if SelArt <> Nil then SelArt.free ; // Sinon appui sur bouton arrêt -> mémoire non libérée !

if TobErreur <> Nil then
    If TobErreur.Detail.Count <> 0 then
        begin
        TheTob:=TobErreur;
{$IFDEF BTP}
        AGLLanceFiche('BTP','BTAFFICHEERREUR','','','') ;
{$ELSE}
        AGLLanceFiche('GC','GCAFFICHEERREUR','','','') ;
{$ENDIF}
        end
        else
        HShowMessage('2;'+TexteMessage[1]+';'+TexteMessage[2]+';I;O;O;N;','','');
if TobErreur <> Nil then TobErreur.Free;
LaTOB.Free ;
LaTob:=Nil ;
if F.bSelectAll.Down then
    begin
    F.bSelectAllClick(Nil);
    F.bSelectAll.Down := False;
    end;
F.FListe.ClearSelected;
F.BChercheClick(Nil);
//FiniMove;
End;

//////////////////// Traitement de la selection ////////////////////////
Procedure TOF_GCMULSUPPRART.TraiteSelection ;
var
   iInd : integer;
   TobArt : TOB;
   LastErr, LastErrMsg, ListeErreur :string;
   SupArt : T_SupArt ;
//   io : TIOErr ;
begin
PB1.Position := 0;
PB1.Step := Trunc(100 / SelArt.Detail.Count);
TWSUPPR.Show;
TobArt := TOB.Create('ARTICLE',NIL,-1);

BeginTrans;

try ExecuteSQL('create index BLO_CLESUPPR on LIGNEOUV(BLO_ARTICLE)'); except; end;

ArreterClick := False;
for iInd := 0 to SelArt.Detail.Count - 1 do
   begin
   Application.ProcessMessages;
   if ArreterClick then Break;
   TobSupp := SelArt.Detail[iInd];
   TobArt.PutValue('GA_ARTICLE', TobSupp.GetValue('GA_ARTICLE'));
   TobArt.LoadDB;
   Art.Caption := TobSupp.GetValue('GA_ARTICLE');
//   ListeErreur:=VerificationArticle(TobSupp.GetValue('GA_ARTICLE'),TobSupp.GetValue('GA_CODEARTICLE'));
   ListeErreur := ControlArticle(TobSupp.GetValue('GA_ARTICLE'),
                                 TobSupp.GetValue('GA_CODEARTICLE'),
                                 (TobArt.GetValue('GA_STATUTART')='DIM'),
                                 TobArt.GetValue('GA_TYPEARTICLE')) ;

   if ListeERreur = '' then
       begin
       SupArt         := T_SupArt.Create;
       SupArt.Cle     := TobSupp.GetValue('GA_ARTICLE') ;
       SupArt.Code    := TobSupp.GetValue('GA_CODEARTICLE');
       SupArt.TypeArt := TobSupp.GetValue('GA_TYPEARTICLE');

       // DCA - FQ MODE 10815
       SupArt.Statut := TobSupp.GetValue('GA_STATUTART');
       try
        SupArt.SuppressionArticle;
       except
        ListeErreur := '1;Erreur inconnue' ;
        ArreterClick := True;
       end;
        //if io <> oeOk then ListeErreur := '1;Erreur inconnue' ;
       SupArt.Destroy;
       end;
   while ListeErreur <> '' do
       Begin
       LastErr := ReadTokenSt(ListeErreur);
       LastErrMsg := ReadTokenSt(ListeErreur);
       MessageErreur(TobSupp.GetValue('GA_ARTICLE'),
                     TobSupp.GetValue('GA_CODEARTICLE'),
                     TobArt.GetValue('GA_LIBELLE'), LastErrMsg, TobErreur, StrToInt(LastErr)) ;
       End ;
   PB1.StepIt;
   end;

if ArreterClick then
    begin
    TobErreur.Detail.Clear;
    TobErreur := nil;
    end;

TobArt.Free;

try ExecuteSQL('drop index LIGNEOUV.BLO_CLESUPPR'); except; end;

if not ArreterClick then CommitTrans else RollBack;

TobSupp.Free;

TWSUPPR.Hide;

end;

//////////// Vérification d'un article pour la suppression //////////
(* Function TOF_GCMULSUPPRART.VerificationArticle(Article,CodeArticle: String): String  ;
Var erreur :string ;
io : TIOErr ;
begin
erreur:=ControlArticle(Article,CodeArticle) ;
if erreur<>'' then
   begin
   //LastError:=StrToInt(ReadTokenSt(erreur));
   //LastErrorMsg:=ReadTokenSt(erreur);
   Result:=erreur ;
   end
Else
    Begin
    io:=Transactions(SuppressionArticle,2);
    if io <>oeOk then
      begin
      //LastError:=1 ;
      //LastErrorMsg:='Erreur inconnue' ;
      Result:='1;Erreur inconnue' ;
      end;
    End;
end;

//////////////////// Suppression d'un article ////////////////////////
Procedure TOF_GCMULSUPPRART.SuppressionArticle ;
Var Cle,Code: String ;
begin
Cle:=TobSupp.GetValue('GA_ARTICLE');
Code:=TobSupp.GetValue('GA_CODEARTICLE');
if (ctxMode in V_PGI.PGIContexte) then
   ExecuteSQL('delete from tarif where gf_article in (select ga_article from article where ga_codearticle="'+Code+'")')
   else ExecuteSQL('DELETE FROM TARIF WHERE GF_ARTICLE="'+Cle+'"');
ExecuteSQL('DELETE FROM ARTICLE WHERE GA_CODEARTICLE="'+Code+'"');
ExecuteSQL('DELETE FROM CONDITIONNEMENT WHERE GCO_ARTICLE="'+Code+'"');
ExecuteSQL('DELETE FROM ARTICLELIE WHERE GAL_ARTICLE="'+Code+'"');
ExecuteSQL('DELETE FROM DISPO WHERE GQ_ARTICLE="'+Code+'"');
ExecuteSQL('DELETE FROM ARTICLETIERS WHERE GAT_ARTICLE="'+Code+'"');
ExecuteSQL('DELETE FROM NOMENENT WHERE GNE_ARTICLE="'+Cle+'"');
ExecuteSQL('DELETE FROM ACTIVITEPAIE WHERE ACP_ARTICLE="'+Code+'"');
ExecuteSQL('DELETE FROM RESSOURCEPR WHERE ARP_ARTICLE="'+Code+'"');
ExecuteSQL('DELETE FROM COMMISSION WHERE GCM_ARTICLE="'+Cle+'"');
ExecuteSQL('UPDATE CATALOGU SET GCA_ARTICLE="" WHERE GCA_ARTICLE="'+Cle+'"') ;
//ExecuteSQL('DELETE FROM NOMENLIG WHERE GNL_CODEARTICLE="'+Code+'"');
ExecuteSQL('DELETE FROM LIENSOLE WHERE LO_IDENTIFIANT="'+Cle+'"');
ExecuteSQL('DELETE FROM GTRADARTICLE WHERE GTA_ARTICLE="'+Cle+'"');
ExecuteSQL('DELETE FROM ARTICLEPIECE WHERE GAP_ARTICLE="'+Cle+'"');
end; *)

Procedure AFLanceFiche_Mul_Suppr_Art;
begin
AGLLanceFiche ('AFF','AFMULSUPPRART','','','');
end;

Initialization
Registerclasses([TOF_GCMULSUPPRART]);
RegisterAglProc('SupprimeArticle',TRUE,1,AGLSupprimeArticle);
end.

