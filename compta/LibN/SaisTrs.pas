unit SaisTrs;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  FichList, Mask, Hctrls, StdCtrls, HSysMenu, hmsgbox, Db, DBTables, Hqry,
  HTB97, DBCtrls, ExtCtrls, HPanel, Grids, DBGrids, HDB, HEnt1, UiUtil, Ent1,
  ZCompte ;

procedure SaisieTrs(Quel : string ; Comment : TActionFiche) ;

const MAXLIGNES           = 2 ;

const RC_LIBGUIDE         = 0 ;
      RC_DOUBLON          = 1 ;
      RC_NONVALIDE        = 2 ;
      RC_CONFIRMDEL       = 3 ;
      RC_BADCOMPTE        = 4 ;
      
type
  TFSaisieTrs = class(TFFicheListe)
    TNaturePiece: THLabel;
    HLabel2: THLabel;
    HLabel3: THLabel;
    HLabel4: THLabel;
    HLabel5: THLabel;
    FCompte1: THCritMaskEdit;
    FCompte2: THCritMaskEdit;
    FLib1: TEdit;
    FLib2: TEdit;
    FTva1: THValComboBox;
    FTva2: THValComboBox;
    HLabel1: THLabel;
    FCode: THDBEdit;
    FLibelle: THDBEdit;
    FNaturePiece: THDBValComboBox;
    FRefReleve: THDBEdit;
    TaGU_TYPE: TStringField;
    TaGU_GUIDE: TStringField;
    TaGU_LIBELLE: TStringField;
    TaGU_ABREGE: TStringField;
    TaGU_JOURNAL: TStringField;
    TaGU_NATUREPIECE: TStringField;
    TaGU_DEVISE: TStringField;
    TaGU_TYPECTREPARTIE: TStringField;
    TaGU_ETABLISSEMENT: TStringField;
    TaGU_UTILISATEUR: TStringField;
    TaGU_DATECREATION: TDateTimeField;
    TaGU_DATEMODIF: TDateTimeField;
    TaGU_SOCIETE: TStringField;
    TaGU_SAISIEEURO: TStringField;
    TaGU_REFRELEVE: TStringField;
    procedure Bourrage(Sender: TObject);
    procedure Modif(Sender: TObject);
    procedure FListeEnter(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BinsertClick(Sender: TObject);
    procedure bDefaireClick(Sender: TObject);
    procedure TestBourrage(Sender: TObject);
  private
    bModif  : Boolean ;
    ECompte : string ;
    function  GetGuideCode : string ;
    // Fonctions ressource
    function  GetMessageRC(MessageID : Integer) : string ;
    function  PrintMessageRC(MessageID : Integer) : Integer ;
    function  Save : Boolean ;
    procedure EnableButtons ;
  public
    procedure NewEnreg ; override ;
    procedure ChargeEnreg ; override ;
    function  EnregOK : Boolean ; override ;
    function  OnDelete : Boolean ; override ;
  end;

implementation

{$R *.DFM}

//=======================================================
//====== Point d'entrée dans la saisie transaction ======
//=======================================================
procedure SaisieTrs(Quel : string ; Comment : TActionFiche) ;
var Transaction : TFSaisieTrs ; PP : THPanel ;
begin
Transaction:=TFSaisieTrs.Create(Application) ;
Transaction.InitFL('GU', 'PRT_GUIDE', Quel, '', Comment, TRUE, Transaction.TaGU_GUIDE,
                   Transaction.TaGU_LIBELLE, nil, ['TTGUIDERLV']) ;
PP:=FindInsidePanel ;
if PP=nil then
  begin
    try
      Transaction.ShowModal ;
    finally
      Transaction.Free ;
    end ;
  end else
  begin
  InitInside(Transaction, PP) ;
  Transaction.Show ;
  end ;
Screen.Cursor:=SyncrDefault ;
end ;

//=======================================================
//================ Evénements de la Form ================
//=======================================================
procedure TFSaisieTrs.FormShow(Sender: TObject);
begin
inherited;
//FTva1.Items[0]:=TraduireMemoire('<<Aucun>>') ;
//FTva2.Items[0]:=TraduireMemoire('<<Aucun>>') ;
ECompte:='' ;
EnableButtons ;
end;

procedure TFSaisieTrs.EnableButtons ;
begin
if FCompte1.Text='' then
  begin
  FCompte2.Enabled:=FALSE ; FCompte2.Color:=clBtnFace ;
  FLib2.Enabled:=FALSE ;    FLib2.Color:=clBtnFace ;
  end else
  begin
  FCompte2.Enabled:=TRUE ;  FCompte2.Color:=clWindow ;
  FLib2.Enabled:=TRUE ;     FLib2.Color:=clWindow ;
  end ;
end ;

//=======================================================
//================= Fonctions Ressource =================
//=======================================================
function TFSaisieTrs.GetMessageRC(MessageID : Integer) : string ;
begin
Result:=HM2.Mess[MessageID] ;
end ;

function TFSaisieTrs.PrintMessageRC(MessageID : Integer) : Integer ;
begin
Result:=HM2.Execute(MessageID, Caption, '') ;
end ;

//=======================================================
//=============== Fonctions dérivées Form ===============
//=======================================================
procedure TFSaisieTrs.NewEnreg ;
begin
inherited ;
FCode.Text:=GetGuideCode ; taGU_GUIDE.AsString:=FCode.Text ;
FLibelle.SetFocus ;
taGU_TYPE.AsString:='RLV' ;
taGU_ETABLISSEMENT.AsString:=VH^.EtablisDefaut ;
taGU_UTILISATEUR.AsString:=V_PGI.User ;
//if bNewObj then Q.FindField('GU_DATECREATION').AsDateTime:=Date ;
//Q.FindField('GU_DATEMODIF').AsDateTime:=Date ;
taGU_SOCIETE.AsString:=V_PGI.CodeSociete ;
taGU_SAISIEEURO.AsString:='-' ;
FCompte1.Text:='' ; FLib1.Text:='' ; FTva1.ItemIndex:=-1 ;
FCompte2.Text:='' ; FLib2.Text:='' ; FTva2.ItemIndex:=-1 ;
bModif:=FALSE ;
end ;

procedure TFSaisieTrs.ChargeEnreg ;
var Q : TQuery ; em : THCritMaskEdit ; ed : TEdit ; cb : THValComboBox ;
    i : Integer ;
begin
Save ;
inherited ;
// Lignes du guide (max=2)
Q:=OpenSQL('SELECT * FROM ECRGUI WHERE EG_TYPE="RLV" AND EG_GUIDE="'+FCode.Text+'" ORDER BY EG_GUIDE, EG_NUMLIGNE', TRUE) ;
for i:=1 to MAXLIGNES do begin
  if Q.EOF then break ;
  em:=THCritMaskEdit(FindComponent('FCompte'+IntToStr(i))) ;
  if em<>nil then em.Text:=Q.FindField('EG_GENERAL').AsString ;
  ed:=TEdit(FindComponent('FLib'+IntToStr(i))) ;
  if ed<>nil then ed.Text:=Q.FindField('EG_LIBELLE').AsString ;
  cb:=THValComboBox(FindComponent('FTva'+IntToStr(i))) ;
  if cb<>nil then cb.Value:=Q.FindField('EG_TVA').AsString ;
  Q.Next ;
  end ;
Ferme(Q) ;
bModif:=FALSE ;
end ;

function TFSaisieTrs.Save : Boolean ;
var Rep : Integer ;
begin
Result:=FALSE ;
if bModif then
   begin
   if FAutoSave.Checked then Rep:=mrYes else Rep:=HM.execute(0,Caption,'') ;
   end else Rep:=321 ;
   case rep of
        mrYes : if not EnregOK then Exit ;
        mrNo  : Exit ;
     mrCancel : Exit ;
   end ;
Result:=TRUE  ;
end ;

function TFSaisieTrs.EnregOK : Boolean ;
var Q : TQuery ; em : THCritMaskEdit ; ed : TEdit ; cb : THValComboBox ;
    i : Integer ;
begin
Result:=inherited EnregOK  ;
//if not Result then Exit ;
ExecuteSQL('DELETE FROM ECRGUI WHERE EG_TYPE="RLV" AND EG_GUIDE="'+FCode.Text+'"') ;
ExecuteSQL('DELETE FROM ANAGUI WHERE AG_TYPE="RLV" AND AG_GUIDE="'+FCode.Text+'"') ;
// Lignes du guide (max=2)
for i:=1 to MAXLIGNES do begin
  Q:=OpenSQL('SELECT * FROM ECRGUI WHERE EG_TYPE="RLV" AND EG_GUIDE="'+W_W+'"', FALSE) ;
  Q.Insert ; InitNew(Q) ;
  Q.FindField('EG_TYPE').AsString:='RLV' ;
  Q.FindField('EG_GUIDE').AsString:=FCode.Text ;
  Q.FindField('EG_NUMLIGNE').AsInteger:=i ;
  em:=THCritMaskEdit(FindComponent('FCompte'+IntToStr(i))) ;
  if em<>nil then Q.FindField('EG_GENERAL').AsString:=em.Text ;
  Q.FindField('EG_AUXILIAIRE').AsString:='' ;
  Q.FindField('EG_REFINTERNE').AsString:='' ;
  ed:=TEdit(FindComponent('FLib'+IntToStr(i))) ;
  if ed<>nil then Q.FindField('EG_LIBELLE').AsString:=ed.Text ;
  Q.FindField('EG_DEBITDEV').AsString:='' ;
  Q.FindField('EG_CREDITDEV').AsString:='' ;
  Q.FindField('EG_MODEPAIE').AsString:='' ;
  Q.FindField('EG_DATEECHEANCE').AsString:='' ;
  Q.FindField('EG_REFEXTERNE').AsString:='' ;
  Q.FindField('EG_DATEREFEXTERNE').AsString:='' ;
  Q.FindField('EG_REFLIBRE').AsString:='' ;
  Q.FindField('EG_AFFAIRE').AsString:='' ;
  Q.FindField('EG_QTE1').AsString:='' ;
  Q.FindField('EG_QTE2').AsString:='' ;
  Q.FindField('EG_QUALIFQTE1').AsString:='' ;
  Q.FindField('EG_QUALIFQTE2').AsString:='' ;
  Q.FindField('EG_MODEREGLE').AsString:='' ;
//  Q.FindField('EG_ECHEANCES').AsString:='' ;
  Q.FindField('EG_ARRET').AsString:='' ;
  Q.FindField('EG_TVAENCAIS').AsString:='-' ;
  cb:=THValComboBox(FindComponent('FTva'+IntToStr(i))) ;
  if cb<>nil then Q.FindField('EG_TVA').AsString:=cb.Value ;
  Q.FindField('EG_RIB').AsString:='-' ;
  Q.FindField('EG_BANQUEPREVI').AsString:='-' ;
  Q.Post ;
  Ferme(Q) ;
  end ;
bModif:=FALSE ;
end ;

function TFSaisieTrs.OnDelete : Boolean ;
begin
Result:=inherited OnDelete  ; if not Result then Exit ;
//ExecuteSQL('DELETE FROM GUIDE WHERE GU_TYPE="RLV" AND GU_GUIDE="'+FCode.Text+'"') ;
ExecuteSQL('DELETE FROM ECRGUI WHERE EG_TYPE="RLV" AND EG_GUIDE="'+FCode.Text+'"') ;
ExecuteSQL('DELETE FROM ANAGUI WHERE AG_TYPE="RLV" AND AG_GUIDE="'+FCode.Text+'"') ;
FCompte1.Text:='' ; FLib1.Text:='' ; FTva1.ItemIndex:=-1 ;
FCompte2.Text:='' ; FLib2.Text:='' ;
bModif:=FALSE ;
EnableButtons ;
end ;

//=======================================================
//================= Fonctions évènements ================
//=======================================================
procedure TFSaisieTrs.Modif(Sender: TObject);
begin
inherited;
bModif:=TRUE ;
end;

procedure TFSaisieTrs.FListeEnter(Sender: TObject);
begin
inherited;
Save ;
end;

//=======================================================
//============== Fonctions utilitaires SLQ ==============
//=======================================================
function TFSaisieTrs.GetGuideCode : string ;
var Q : TQuery ; sNum : string ; iNum : Integer ;
begin
Q:=OpenSQL('SELECT MAX(GU_GUIDE) FROM GUIDE WHERE GU_TYPE="RLV"', TRUE) ;
sNum:='001' ;
if not Q.EOF then
   begin
   sNum:=Q.Fields[0].AsString ;
   if sNum<>'' then
      begin
      iNum:=Trunc(Valeur(sNum)) ;
      sNum:=Format('%.3d', [iNum+1]);
      end else sNum:='001';
   end ;
Ferme(Q) ;
Result:=sNum ;
end ;

procedure TFSaisieTrs.TestBourrage(Sender: TObject);
begin
inherited;
ECompte:=THCritMaskEdit(Sender).Text ;
end;

procedure TFSaisieTrs.Bourrage(Sender: TObject);
var Comptes : TZCompte ; ed : TEdit ; cb : THValComboBox ;
    NumCompte, sName : string ; k : Integer ;
begin
inherited;
NumCompte:=THCritMaskEdit(Sender).Text ;
if ECompte=NumCompte then Exit ;
Comptes:=TZCompte.Create ;
k:=Comptes.GetCompte(NumCompte) ;
if k>=0 then
   begin
   sName:=THCritMaskEdit(Sender).Name ; Delete(sName, 1, Length('FCompte')) ;
   ed:=TEdit(FindComponent('FLib'+sName)) ;
   if ed<>nil then ed.Text:=Comptes.GetValue('G_LIBELLE', k) ;
   cb:=THValComboBox(FindComponent('FTva'+sName)) ;
   if cb<>nil then
     if Comptes.GetValue('G_TVA', k)<>'' then cb.Value:=Comptes.GetValue('G_TVA', k)
                                         else cb.ItemIndex:=0 ;
   end ;
Comptes.Free ;
THCritMaskEdit(Sender).Text:=NumCompte ;
if (k<0) and (THCritMaskEdit(Sender).Name<>'FCompte2') then
  begin PrintMessageRC(RC_BADCOMPTE) ; THCritMaskEdit(Sender).SetFocus ; end ;
EnableButtons ;
end;

procedure TFSaisieTrs.BinsertClick(Sender: TObject);
begin
inherited;
EnableButtons ;
end;

procedure TFSaisieTrs.bDefaireClick(Sender: TObject);
begin
inherited;
EnableButtons ;
end;

end.
