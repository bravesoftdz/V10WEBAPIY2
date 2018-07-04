unit EtapeReg;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  FichList, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  StdCtrls, Hcompte, HDB,Hctrls, DBCtrls, Mask,
  hmsgbox, Buttons, ExtCtrls, Grids, DBGrids, Ent1, HEnt1, EncUTil,
  HSysMenu, MajTable, HTB97, Hqry, HRegCpte, HPanel, UiUtil, ComCtrls,
  ADODB ;

Procedure ParamEtapeReg (Enc: Boolean; lequel : String ; Bloc : boolean ) ;
function isOracle : boolean;

type TGuide = RECORD
  Code,Lib,dev :String ;
  Cpte1,Cpte2 : String17 ;
  END ;

type
  TFEtapeReg = class(TFFicheListe)
    HLabel3: THLabel;
    ER_ETAPE: TDBEdit;
    ENCDEC: TDBCheckBox;
    HLabel5: THLabel;
    ER_LIBELLE: TDBEdit;
    HLabel8: THLabel;
    ER_TYPEETAPE: THDBValComboBox;
    ER_GLOBALISE: TDBCheckBox;
    GComptes: TGroupBox;
    HLabel4: THLabel;
    HLabel9: THLabel;
    EDEPART: THLabel;
    EARRIVEE: THLabel;
    ER_CPTEARRIVEE: THDBCpteEdit;
    ER_CPTEDEPART: THDBCpteEdit;
    TaER_ENCAISSEMENT: TStringField;
    TaER_ETAPE: TStringField;
    TaER_LIBELLE: TStringField;
    TaER_TYPEETAPE: TStringField;
    TaER_CPTEDEPART: TStringField;
    TaER_CPTEARRIVEE: TStringField;
    TaER_CATEGORIEMP: TStringField;
    TaER_MODEPAIE: TStringField;
    TaER_DEVISE: TStringField;
    TaER_GUIDE: TStringField;
    TaER_GLOBALISE: TStringField;
    QMP: TQuery;
    TaER_EXPORTCFONB: TStringField;
    TaER_FORMATCFONB: TStringField;
    TaER_BORDEREAU: TStringField;
    TaER_DOCUMENT: TStringField;
    TaER_ENVOITRANS: TStringField;
    Pages: TPageControl;
    TS1: TTabSheet;
    TS2: TTabSheet;
    ER_EXPORTCFONB: TDBCheckBox;
    HLabel10: THLabel;
    ER_FORMATCFONB: THDBValComboBox;
    ER_BORDEREAU: TDBCheckBox;
    TER_DOCUMENT: THLabel;
    ER_DOCUMENT: THDBValComboBox;
    HLabel1: THLabel;
    LIBGUIDE: TEdit;
    BZOOMETAPE: TToolbarButton97;
    HLabel6: THLabel;
    ER_DEVISE: THDBValComboBox;
    ER_GUIDE: TDBEdit;
    TER_CODEAFB: THLabel;
    ER_CODEAFB: THDBMultiValComboBox;
    HLabel2: THLabel;
    ER_MODEPAIE: THDBValComboBox;
    HLabel7: THLabel;
    ER_CATEGORIEMP: THDBValComboBox;
    ER_ENVOITRANS: THDBValComboBox;
    TER_ENVOITRANS: THLabel;
    TaER_CODEAFB_SQL: TMemoField;
    TaER_CODEAFB_ORA: TStringField;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ER_TYPEETAPEExit(Sender: TObject);
    procedure BZOOMETAPEClick(Sender: TObject);
    procedure ER_CPTEARRIVEEExit(Sender: TObject);
    procedure STaDataChange(Sender: TObject; Field: TField);
    procedure ER_CPTEDEPARTExit(Sender: TObject);
    procedure ER_MODEPAIEChange(Sender: TObject);
    procedure ER_CATEGORIEMPExit(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure ER_LIBELLEExit(Sender: TObject);
    procedure LIBGUIDEKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure LIBGUIDEExit(Sender: TObject);
    procedure ER_EXPORTCFONBClick(Sender: TObject);
    procedure ER_BORDEREAUClick(Sender: TObject);
  private    { Déclarations privées }
    FGuide        : TGuide ;
    Enc,OkGuideVu : Boolean ;
    TypGuide           : String3 ;
    Bloc          : Boolean ;
    Function  ControleComptes : Boolean ;
    Procedure ChargeComptes ;
    Procedure TrouvGeneRest ;
    function  OkGuide(OkMss : boolean) : Boolean ;
    function  TrouveGuide : boolean ;
    Function  CreerGuide : boolean ;
    Procedure DetruitAutresGuides ;
    procedure ChargeEnreg1 ;
    procedure CATTOMPENCDEC ( CAT : String3 ; it,va : HTStrings ) ;
    procedure DetruitLeGuide(Guide : String3) ;
  public    { Déclarations publiques }
    Function  EnregOK : boolean ; Override ;
    Procedure NewEnreg ; Override ;
  end;

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  CPGUIDE_TOM,UtilPgi ;

{$R *.DFM}


Procedure ParamEtapeReg (Enc: Boolean ; Lequel : String ; Bloc : boolean ) ;
var FETAPEREGLE: TFEtapeReg;
    StEnc : String ;
    PP : THPanel ;
BEGIN
if Bloc then if _Blocage(['nrBatch','nrCloture'],True,'nrBatch') then Exit ;
FETAPEREGLE:=TFEtapeReg.Create(Application) ;
FEtapeRegle.Bloc:=Bloc ;
FEtapeRegle.Enc:=Enc ; if ENC then StEnc:='X' else StEnc:='-' ;
if Enc then FEtapeRegle.HelpContext := 1480000 else FEtapeRegle.HelpContext := 1490000 ;
FETAPEREGLE.InitFL('ER','PRT_ETAPEREG',LeQuel,StEnc,taModif,TRUE,FETAPEREGLE.TaER_ETAPE,
             FETAPEREGLE.TaER_LIBELLE,nil,
             ['ttEtapeEncais','ttEtapeDecais','ttEtapeDecRegle','ttEtapeDecCheque','ttEtapeDecTraite','ttEtapeEncTraite']) ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
   try
     FETAPEREGLE.ShowModal ;
     finally
     FETAPEREGLE.Free ;
     if Bloc then _Bloqueur('nrBatch',False) ;
     END ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FETAPEREGLE,PP) ;
   FETAPEREGLE.Show ;
   END ;
END ;

procedure TFEtapeReg.FormCreate(Sender: TObject);
begin
  { FQ 21463 BVE 24.09.07 : Probleme ER_CODEAFB est un LongChar qui est transformé en Memo sous MsSQL
  et en VarChar(2000) sous Oracle on ne peut donc pas les gérer exactement de la meme facon. }
  if isOracle then
  begin
     TaER_CODEAFB_SQL.Free;
     TaER_CODEAFB_ORA.FieldName := 'ER_CODEAFB';
  end
  else
  begin
     TaER_CODEAFB_ORA.Free;
     TaER_CODEAFB_SQL.FieldName := 'ER_CODEAFB';
  end;
  { END FQ 21463 }
  inherited;
end;

procedure TFEtapeReg.FormShow(Sender: TObject);
var KENC : String1 ;
BEGIN

  {SG6 18.04.05 FQ 15372 / JP 08/08/05 : + FQ 15758 ???}
  ER_CATEGORIEMP.Plus := ' AND CO_CODE<>"TRI"';
  {JP 05/08/05 : FQ 15168 : On interdit les virements internationaux dans le suivi des tiers}
  ER_FORMATCFONB.Plus := ' AND CO_CODE <> "VRI"';

If Not VH^.OldTeleTrans Then BEGIN ER_ENVOITRANS.Visible:=FALSE ; TER_ENVOITRANS.Visible:=FALSE ; END ;
if not isOracle then
   ChangeSizeLongChar(taER_CODEAFB_SQL);
if ENC then BEGIN Caption:=HM2.Mess[0] ; TypGuide:='ENC' ; KENC:='X' ; END
       else BEGIN Caption:=HM2.Mess[1] ; TypGuide:='DEC' ; KEnc:='-' ; END ;
 inherited;
if (not (Ta.Eof)) then ER_ETAPE.Enabled:=False ;
if(Ta.Eof) And (Ta.Bof) And (FTypeAction<>taConsult) then
   BEGIN
   if Ta.State=dsInsert then NewEnreg else BinsertClick(Nil) ;
   END ;
UpdateCaption(Self) ;
end;

procedure TFEtapeReg.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Parent is THPanel then if Bloc then _Bloqueur('nrBatch',False) ;
DetruitAutresGuides ;
 inherited;
end;

Procedure TFEtapeReg.NewEnreg ;
BEGIN
Inherited ;
TaER_GLOBALISE.Value:='-' ;
TaER_GUIDE.Value:='' ; LIBGUIDE.Text:='' ;
if Enc then TaER_ENCAISSEMENT.Value:='X' else TaER_ENCAISSEMENT.Value:='-' ;
END ;

Function TFEtapeReg.EnregOK : boolean ;
var Q : TQuery ;
BEGIN
result:=InHerited EnregOK ; if Not Result then Exit ;
Modifier:=True ;
if ((Result) and (Ta.state in [dsEdit,dsInsert])) then
   BEGIN
   Result:=FALSE ;
   if Ta.State=dsInsert then
     BEGIN
     if Presence('ETAPEREG','ER_ETAPE',ER_ETAPE.Text) then
        BEGIN
        HM.Execute(4,'','') ; if ER_ETAPE.CanFocus then ER_ETAPE.SetFocus ;
        Exit ;
        END ;
     END ;
   if TaER_TYPEETAPE.AsString='' then
      BEGIN
      HM2.Execute(15,'','') ;
      if ER_TYPEETAPE.CanFocus then ER_TYPEETAPE.SetFocus ;
      Exit ;
      END ;
   if TaER_CPTEDEPART.AsString='' then
      BEGIN
      HM2.Execute(16,'','') ;
      if ER_CPTEDEPART.CanFocus then ER_CPTEDEPART.SetFocus ;
      Exit ;
      END ;
   if ER_CPTEDEPART.ExisteH <=0 then
      BEGIN
      HM2.Execute(21,'','') ;
      if ER_CPTEDEPART.CanFocus then ER_CPTEDEPART.SetFocus ;
      Exit ;
      END ;
   if ((TaER_CPTEARRIVEE.AsString='') and (TaER_TYPEETAPE.AsString<>'POR')) then
     BEGIN
     HM2.Execute(17,'','') ;
     if ER_CPTEARRIVEE.CanFocus then ER_CPTEARRIVEE.SetFocus ; Exit ;
     END Else
     BEGIN
     if ER_CPTEARRIVEE.Enabled then
       BEGIN
       if ER_CPTEARRIVEE.ExisteH <=0 then
          BEGIN
          HM2.Execute(22,'','') ;
          if ER_CPTEARRIVEE.CanFocus then ER_CPTEARRIVEE.SetFocus ;
          Exit ;
          END ;
       END ;
     END ;
   if TaER_CATEGORIEMP.AsString='' then
      BEGIN
      HM2.Execute(18,'','') ;
      if ER_CATEGORIEMP.CanFocus then ER_CATEGORIEMP.SetFocus ;
      Exit ;
      END ;
   if TaER_DEVISE.AsString='' then
     BEGIN
     HM2.Execute(19,'','') ;
     if ER_DEVISE.CanFocus then ER_DEVISE.SetFocus ;
     Exit ;
     END else
     BEGIN
     if (ER_TYPEETAPE.Value='BQE') then
       BEGIN
       Q:=OpenSQL('Select BQ_DEVISE from BANQUECP Where BQ_GENERAL="'+ER_CPTEARRIVEE.Text
                 +'" AND BQ_NODOSSIER="'+V_PGI.NoDossier+'"',True) ; // 19/10/2006 YMO Multisociétés
       if Not Q.EOF then if (ER_DEVISE.Value<>Q.Fields[0].AsString) then
         BEGIN
         HM2.Execute(23,'','') ;
         if ER_DEVISE.CanFocus then ER_DEVISE.SetFocus ;
         Ferme(Q) ;
         Exit ;
         END ;
       Ferme(Q) ;
       END ;
     if (ER_DEVISE.Value<>V_PGI.DevisePivot) and ER_EXPORTCFONB.Checked then
       BEGIN
       HM2.Execute(24,'','') ;
       if ER_DEVISE.CanFocus then ER_DEVISE.SetFocus ;
       Exit ;
       END ;
     END ;
   if not ControleComptes then Exit ;
   if not TrouveGuide then
      BEGIN
      if Not CreerGuide then Exit ;
      END else
      BEGIN
      if Not OkGuide(False) then
         BEGIN
         DetruitLeGuide(ER_GUIDE.Text) ;
         if Not CreerGuide then Exit ;
         END ;
      END ;
   if (ta.State=dsInsert) and not OkGuideVu then
     BEGIN
     OkGuideVu:=True ;
     if HM2.Execute(8,'','')=mrYes then BZOOMETAPEClick(nil) ;
     END ;
   END ;
result:=TRUE  ; Modifier:=False ;
END ;

Function TFEtapeReg.ControleComptes : Boolean ;
Var NumErr: Integer ;
    Q     : TQuery ;
BEGIN
NumErr:=0 ;
if (ER_CPTEDEPART.Text=ER_CPTEARRIVEE.Text) then NumErr:=20 ;
if not BonCompteEnc (Enc , ER_CPTEDEPART.Text) then NumErr:=21 ;
if ER_CPTEARRIVEE.Text <>'' then
  BEGIN
  if ER_TYPEETAPE.Value<>'BQE' then if not BonCompteEnc (Enc , ER_CPTEARRIVEE.Text) then
    BEGIN
    NumErr:=22 ;
    END else
    BEGIN
    Q:=OpenSQL('Select G_GENERAL from Generaux Where G_GENERAL="'+ER_CPTEARRIVEE.Text+'"',True) ;
    if Q.Eof then NumErr:=22 ;
    Ferme(Q) ;
    END ;
  END ;
if NumErr <>0 then HM2.Execute(NumErr,'','') ;
Result:=(NumErr=0) ;
END ;

Procedure TFEtapeREG.ChargeComptes ;
BEGIN
if Enc then
  BEGIN
  ER_CPTEDEPART.ZoomTable:=tzGEncais ; ER_CPTEARRIVEE.ZoomTable:=tzGBQCaissCli ;
  END else
  BEGIN
  ER_CPTEDEPART.ZoomTable:=tzGDecais ; ER_CPTEARRIVEE.ZoomTable:=tzGBQCaissFou ;
  END ;
if (ER_TYPEETAPE.Value='BQE') then ER_CPTEARRIVEE.ZoomTable:=TzGBanque ;
END ;

procedure TFEtapeReg.ER_TYPEETAPEExit(Sender: TObject);
begin
  inherited;
ChargeComptes ;
ER_CPTEDEPART.ExisteH ;
if ER_TYPEETAPE.Value='POR' then
  BEGIN
  if ER_MODEPAIE.ItemIndex<=0 then
    BEGIN
    ER_CPTEARRIVEE.Text:=''; ER_CPTEARRIVEE.Enabled:=False ;
    EARRIVEE.Caption:=HM2.Mess[7] ;
    END ;
  END else
  BEGIN
  ER_CPTEARRIVEE.Enabled:=True ; if ER_CPTEARRIVEE.ExisteH<=0 then BEGIN ER_CPTEARRIVEE.Text:='' ; END ;
  END ;
end;

procedure TFEtapeReg.BZOOMETAPEClick(Sender: TObject);
BEGIN
  inherited;
LibGuideExit(Nil) ; 
OkGuideVu:=True ;
if ta.State in [dsInsert,dsEdit] then if not EnregOk then exit ;
ParamGuide(ER_GUIDE.Text,TypGuide,taModif) ;
Application.ProcessMessages ; TrouveGuide ;
if Not OkGuide(True) then
   BEGIN
   DetruitLeGuide(ER_GUIDE.Text) ; CreerGuide ;
   END ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 08/08/2005
Modifié le ... :   /  /    
Description .. : 
Suite ........ : SBO 08/08/2005 : FQ 15684 : suppression du test limitant 
Suite ........ : l'option "Contrepartie Globalisée"...
Mots clefs ... : 
*****************************************************************}
procedure TFEtapeReg.TrouvGeneRest ;
var Q : TQuery ;
begin
  Q := OpenSQL('Select BQ_BANQUE,BQ_DEVISE from BANQUECP Where BQ_GENERAL="'+ER_CPTEARRIVEE.Text
              +'" AND BQ_NODOSSIER="'+V_PGI.NoDossier+'"',True) ; // 19/10/2006 YMO Multisociétés
  if not Q.EOF then
    if ER_DEVISE.ItemIndex=-1 then
      TaER_DEVISE.AsString := Q.Fields[1].AsString ;
  Ferme(Q) ;
end ;

procedure TFEtapeReg.ER_CPTEARRIVEEExit(Sender: TObject);
begin
  inherited;
ER_CPTEARRIVEE.ExisteH ;
if Ta.State=dsEdit then TrouvGeneRest ;
end;

Function TFEtapeReg.OkGuide(OkMss : boolean) : boolean;
BEGIN
Result:=False ;
if ER_GUIDE.Text<>'' then
  BEGIN
  if ER_DEVISE.Value<>FGuide.Dev then BEGIN if OkMss then HM2.Execute(2,'','') ; exit ; END ;
  if ER_CPTEDEPART.Text<>FGuide.Cpte1 then BEGIN if OkMss then HM2.Execute(3,'','') ; exit ; END ;
  if ER_CPTEARRIVEE.Text<>FGuide.Cpte2 then BEGIN if OkMss then HM2.Execute(4,'','') ; exit ; END ;
  END ;
  //else BEGIN HM2.Execute(5,'','') ; ER_GUIDE.ItemIndex:=0 ; END ;
Result:=True ;
END ;

function TFEtapeReg.TrouveGuide : boolean ;
var Q : TQuery ;
    St: String ;
    numlig : Integer ;
BEGIN
Result:=False ;
With FGuide do
  BEGIN
  Cpte1:='' ; Cpte2:='' ;
  Code:='' ; Lib:='' ; Dev:='' ;
  END ;
if (taER_GUIDE.AsString='') then Exit ;
St:='select gu_guide, gu_libelle, gu_devise, eg_numligne, eg_general from guide left join ecrgui on eg_type=gu_type '+
    'and eg_guide=gu_guide where gu_type="'+TypGuide+'" AND gu_guide="'+ER_GUIDE.Text+'" ORDER BY eg_numligne' ;
Q:=OpenSQL(St,True) ; Result:=not Q.Eof ; //LIBGUIDE.Text:='' ;
While not Q.EOF do
  BEGIN
  numlig:=Q.Fields[3].AsInteger ;
  With FGuide do
    BEGIN
    case numlig of
      1: Cpte1:=Q.Fields[4].AsString ;
      2: BEGIN
         Cpte2:=Q.Fields[4].AsString ;
         Code:=Q.Fields[0].AsString ;
         Lib:=Q.Fields[1].AsString ;
         Dev:=Q.Fields[2].AsString ;
         Application.ProcessMessages ;
         END ;
      END ;
    END ;
  Q.Next ;
  END ;
LIBGUIDE.Text:=FGuide.Lib ;
Ferme(Q) ;
END ;

Function TFEtapeReg.CreerGuide : boolean ;
var TG, TECRG : TQuery ;
    i : integer ;
    LeJournal : String ;

    function FirstJal : String3 ;
    var Q : TQuery ;
        St : String ;
        NatJal : String3 ;
        Trouve : boolean ;
        N : byte ;
    BEGIN
    Result:='' ; N:=0 ; Trouve:=False ; Q:=nil ;
    While (not Trouve) and (N<2) do
      BEGIN
      if (ER_TYPEETAPE.Value='BQE') then NatJal:='BQE' else NatJal:='OD' ;
      St:='SELECT J_JOURNAL FROM JOURNAL WHERE J_NATUREJAL="'+NatJal+'" AND J_JOURNAL<>"'+VH^.JalATP+'" AND J_JOURNAL<>"'+VH^.JalVTP+'"' ;
      if N=0 then
        BEGIN
        if (V_PGI.DevisePivot<>ER_DEVISE.Value) then St:=St+' AND J_MULTIDEVISE="X"' ;
        END ;
      if NatJal='BQE' then St:=St+' AND J_CONTREPARTIE="'+ER_CPTEARRIVEE.Text+'"' ;
      Q:=OpenSQL(St+' ORDER BY J_JOURNAL',True) ;
      Trouve:=Not Q.Eof ;
      if Trouve then Result:=Q.Fields[0].AsString ;
      Ferme(Q) ;
      Inc(N) ;
      if ((Not Trouve) and (V_PGI.DevisePivot=ER_DEVISE.Value)) then Break ;
      END ;
    // Si pas OK, accepter journal d'OD sur les remises en banque
    if ((Result='') and (Not Trouve) and (ER_TYPEETAPE.Value='BQE')) then
       BEGIN
       St:='SELECT J_JOURNAL FROM JOURNAL WHERE J_NATUREJAL="OD" AND J_JOURNAL<>"'+VH^.JalATP+'" AND J_JOURNAL<>"'+VH^.JalVTP+'"' ;
       if V_PGI.DevisePivot<>ER_DEVISE.Value then St:=St+' AND J_MULTIDEVISE="X"' ;
       St:=St+' ORDER BY J_JOURNAL' ;
       Q:=OpenSQL(St,True) ;
       if Not Q.EOF then Result:=Q.Fields[0].AsString ;
       Ferme(Q) ;
       END ;
    END ;

BEGIN
Result:=False ;
LeJournal:=FirstJal ; if LeJournal='' then BEGIN HM2.Execute(25,'','') ; Exit ; END ;
{Entête du guide}
TG:=OpenSQL('Select * from guide where gu_Type="'+TypGuide+'" order by gu_guide',False) ;
TG.UpdateMode:=upWhereChanged ;
if (LIBGUIDE.Text='') then BEGIN LIBGUIDE.Text:=ER_LIBELLE.Text ; Application.ProcessMessages ; END ;
TG.Insert ; InitNew(TG) ;
TG.FindField('GU_TYPE').AsString:=TypGuide ;
TG.FindField('GU_GUIDE').AsString:=ER_ETAPE.Text ;
TG.FindField('GU_LIBELLE').AsString:=LIBGUIDE.Text ;
TG.FindField('GU_JOURNAL').AsString:=LeJournal ;
if Enc then TG.FindField('GU_NATUREPIECE').AsString:='RC' else TG.FindField('GU_NATUREPIECE').AsString:='RF' ;
TG.FindField('GU_DEVISE').AsString:=ER_DEVISE.Value ;
TG.FindField('GU_ETABLISSEMENT').AsString:=VH^.EtablisDefaut ;
TG.FindField('GU_TRESORERIE').AsString:='-' ;
TG.Post ;
Ferme(TG) ;
{Lignes du guide}
TECRG:=OpenSQL('Select * from EcrGui where eg_Type="'+TypGuide+'"',False) ;
TECRG.UpdateMode:=upWhereChanged ;
for i:=1 to 2 do
  BEGIN
  TECRG.Insert ; InitNew(TECRG) ;
  TECRG.FindField('EG_TYPE').AsString:=TypGuide ;
  TECRG.FindField('EG_GUIDE').AsString:=ER_ETAPE.Text ;
  TECRG.FindField('EG_LIBELLE').AsString:=LIBGUIDE.Text ;
  TECRG.FindField('EG_NUMLIGNE').AsInteger:=i ;
  TECRG.FindField('EG_MODEPAIE').AsString:=ER_MODEPAIE.Value ;
  if i=1 then TECRG.FindField('EG_GENERAL').AsString:=ER_CPTEDEPART.Text
         else TECRG.FindField('EG_GENERAL').AsString:=ER_CPTEARRIVEE.Text ;
  TECRG.Post ;
  END ;
Ferme(TECRG) ;
With FGuide do
  BEGIN
  Cpte1:=ER_CPTEDEPART.Text ; Cpte2:=ER_CPTEARRIVEE.Text ;
  Code:=ER_ETAPE.Text       ; Lib:=LIBGUIDE.Text ;
  Dev:=ER_DEVISE.Value      ;
  END ;
if ta.State=dsBrowse then ta.edit ;
taER_GUIDE.AsString:=ER_ETAPE.Text ;
AvertirTable('ttGuideENCDEC') ;
Result:=True ;
END ;

Procedure TFEtapeReg.DetruitAutresGuides ;
var Q : TQuery ;
BEGIN
if ta.State=dsInsert then Q:=OpenSQL('Select GU_GUIDE from GUIDE where GU_TYPE="'+TypGuide+'" AND GU_GUIDE<>"'+ER_ETAPE.Text+'" AND GU_GUIDE NOT IN(SELECT ER_GUIDE FROM ETAPEREG WHERE ER_GUIDE<>"'+ER_ETAPE.text+'")',True)
  else if ER_Etape.Text<>TaER_GUIDE.AsString then Q:=OpenSQL('Select GU_GUIDE from GUIDE where GU_TYPE="'+TypGuide+'" AND GU_GUIDE="'+taER_GUIDE.AsString+'"',True) else exit ;
While not Q.Eof do
  BEGIN
  DetruitLeGuide(Q.Fields[0].AsString) ;
  Q.Next ;
  END ;
Ferme(Q) ;
END ;

procedure TFEtapeReg.DetruitLeGuide(Guide : String3) ;
BEGIN
ExecuteSQL('Delete from GUIDE where gu_type="'+TypGuide+'" and gu_guide="'+Guide+'"') ;
ExecuteSQL('Delete from ECRGUI where eg_type="'+TypGuide+'" and eg_guide="'+Guide+'"') ;
ExecuteSQL('Delete from ANAGUI where ag_type="'+TypGuide+'" and ag_guide="'+Guide+'"') ;
LIBGUIDE.Text:='' ;
END ;

procedure TFEtapeReg.STaDataChange(Sender: TObject; Field: TField);
begin
  inherited;
if Field=nil then ChargeEnreg1 ;
end;

procedure TFEtapeReg.ChargeEnreg1 ;
BEGIN
ER_TYPEETAPEExit(Nil) ;
ER_CATEGORIEMPExit(Nil) ;
if (Trim(TaER_MODEPAIE.AsString)='') then ER_MODEPAIE.ItemIndex:=0 else ER_MODEPAIEchange(Nil) ;
TrouveGuide ;
OkGuideVu:=False ;
END ;

procedure TFEtapeReg.ER_CPTEDEPARTExit(Sender: TObject);
begin
  inherited;
ER_CPTEDEPART.ExisteH ;
end;

procedure TFEtapeReg.ER_MODEPAIEChange(Sender: TObject);
var Q : TQuery;
begin
  inherited;
if (ER_TYPEETAPE.Value='POR') then
  BEGIN
  if (ER_MODEPAIE.ItemIndex>0) then
    BEGIN
    Q:=OpenSQL('SELECT MP_GENERAL FROM MODEPAIE WHERE MP_MODEPAIE="'+ER_MODEPAIE.Value+'"',true) ;
    ER_CPTEARRIVEE.Text:=Trim(Q.Fields[0].AsString) ;
    ER_CPTEARRIVEE.ExisteH ;
    Ferme(Q) ;
    END else
    BEGIN
    ER_CPTEARRIVEE.Text:=''; ER_CPTEARRIVEE.Enabled:=False ;
    EARRIVEE.Caption:=HM2.Mess[7] ;
    END ;
  END ;
end;

procedure TFEtapeReg.ER_CATEGORIEMPExit(Sender: TObject);
var Val : String ;
begin
  inherited;
Val:=taER_MODEPAIE.AsString ;
// BVE 21.09.07 : Probleme MAJ du champ
// CatToMPENCDEC(ER_CATEGORIEMP.Value,ER_MODEPAIE.Items, ER_MODEPAIE.Values ) ;
CatToMPENCDEC(taER_CATEGORIEMP.AsString,ER_MODEPAIE.Items, ER_MODEPAIE.Values ) ;
ER_MODEPAIE.ItemIndex:=ER_MODEPAIE.Values.Indexof(Val) ;
// BVE 21.09.07 : Probleme avec des codes sur 2 caractères.
ER_CATEGORIEMP.ItemIndex := ER_CATEGORIEMP.Values.IndexOf(Trim(taER_CATEGORIEMP.AsString));
if Ta.State=dsEdit then TrouvGeneRest ;
end;

procedure TFEtapeReg.BDeleteClick(Sender: TObject);
begin
if (ER_GUIDE.Text<>'') then DetruitLeGuide(ER_GUIDE.Text) ;
  inherited;
end;

procedure TFEtapeReg.BAnnulerClick(Sender: TObject);
var ok : boolean ;
begin
ok:=OkGuide(False) ;
if not Ok then DetruitLeGuide(ER_GUIDE.Text) ;
  inherited;
if not Ok then CreerGuide ;
end;

procedure TFEtapeReg.CATTOMPENCDEC ( CAT : String3 ; it,va : HTStrings ) ;
BEGIN
It.Clear ; Va.Clear ;
if CAT='' then Exit ;
QMP.SQL.Clear ;
if ENC then
   BEGIN
   QMP.SQL.Add('Select MP_MODEPAIE, MP_LIBELLE from MODEPAIE Where MP_CATEGORIE="'+CAT+'" AND (MP_GENERAL="" ') ;
   QMP.SQL.Add('OR MP_GENERAL IN (SELECT G_GENERAL FROM GENERAUX WHERE ((G_NATUREGENE<>"COF" AND G_NATUREGENE<>"COS" AND G_NATUREGENE<>"TID") OR G_NATUREGENE="BQE" OR G_NATUREGENE="CAI" AND (G_LETTRABLE="X" OR G_COLLECTIF="X"))))') ;
   END else
   BEGIN
   QMP.SQL.Add('Select MP_MODEPAIE, MP_LIBELLE from MODEPAIE Where MP_CATEGORIE="'+CAT+'" AND (MP_GENERAL="" ') ;
   QMP.SQL.Add('OR MP_GENERAL IN (SELECT G_GENERAL FROM GENERAUX WHERE ((G_NATUREGENE<>"COC" AND G_NATUREGENE<>"TIC") OR G_NATUREGENE="BQE" OR G_NATUREGENE="CAI" AND (G_LETTRABLE="X" OR G_COLLECTIF="X"))))') ;
   END ;
It.Add(TraduireMemoire('<<Tous>>')) ; Va.Add('') ;
ChangeSQL(QMP) ; //QMP.Prepare ;
PrepareSQLODBC(QMP) ;
QMP.Open ;
While not QMP.EOF do
  BEGIN
  Va.Add(QMP.Fields[0].AsString) ; It.Add(QMP.Fields[1].AsString) ;
  QMP.Next ;
  END ;
QMP.Close ;
END ;

procedure TFEtapeReg.ER_LIBELLEExit(Sender: TObject);
begin
if (Trim(LIBGUIDE.Text)='') then LIBGUIDE.Text:=ER_LIBELLE.Text ;
end;

procedure TFEtapeReg.LIBGUIDEKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
ta.edit ; ER_GUIDE.Text:=ER_GUIDE.Text ;
end;

procedure TFEtapeReg.LIBGUIDEExit(Sender: TObject);
begin
  inherited;
if (ER_GUIDE.Text='') or (FGuide.Lib='') or (LibGuide.text='') then Exit ;
if (LIBGUIDE.Text<>FGuide.Lib) then
  ExecuteSQL('UPDATE GUIDE SET GU_LIBELLE="'+CheckdblQuote(LibGuide.Text)+'" WHERE GU_TYPE="'+TypGuide+'" AND GU_GUIDE="'+ER_GUIDE.Text+'"') ;
end;

procedure TFEtapeReg.ER_EXPORTCFONBClick(Sender: TObject);
begin
  inherited;
if Not ER_EXPORTCFONB.Checked then
   BEGIN
   { BVE 24.09.07 : Probleme message voulez-vous enregistrez alors qu'aucune modif n'a été faite }
   if Ta.State <> dsBrowse then
      ER_FORMATCFONB.Value:='' ;
   ER_FORMATCFONB.Enabled:=False ;
   END else
   BEGIN
   ER_FORMATCFONB.Enabled:=True ;
   END ;
end;

procedure TFEtapeReg.ER_BORDEREAUClick(Sender: TObject);
begin
  inherited;
if Not ER_BORDEREAU.Checked then
   BEGIN
   ER_DOCUMENT.Value:='' ; ER_DOCUMENT.Enabled:=False ;
   END else
   BEGIN
   ER_DOCUMENT.Enabled:=True ;
   END ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 24/09/2007
Modifié le ... :   /  /    
Description .. : Permet d'indiquer si la base est une version d'Oracle
Suite ........ : FQ 21463 
Mots clefs ... : 
*****************************************************************}
function isOracle : boolean;
begin
  if V_PGI.Driver in [dbOracle7,dbOracle8,dbOracle9,dbOracle10] then
     Result := true
  else
     Result := false;
end;

end.
