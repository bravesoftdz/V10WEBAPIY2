unit FmtImpor;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  FichList,
{$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  DB, Spin, HDB, DBCtrls, StdCtrls, Hctrls, Mask,
  HSysMenu, hmsgbox, Buttons, ExtCtrls, Grids, DBGrids, Ent1, Hent1, Hqry,
  HTB97, HPanel;


// ajout me n'est jamais utilisé function  CodeFormatFixeExiste(FNat : String ; Code : String) : boolean ;
Procedure ParamImport (Nature,Import : String ) ;

type
  TFFmtImport = class(TFFicheListe)
    TIM_CODE: THLabel;
    TIM_LIBELLE: THLabel;
    TIM_FORMATDATE: THLabel;
    TIM_SEPDATE: THLabel;
    TIM_SEPDECIMAL: THLabel;
    TIM_SEPMILLIER: THLabel;
    TIM_DELIMITEUR: THLabel;
    TIM_FORMATSENS: THLabel;
    TLig: THLabel;
    Label1: TLabel;
    HLabel1: THLabel;
    HLabel2: THLabel;
    IM_CODE: TDBEdit;
    IM_LIBELLE: TDBEdit;
    IM_FORMATDATE: THDBValComboBox;
    IM_SEPDATE: THDBValComboBox;
    IM_SEPDECIMAL: THDBValComboBox;
    IM_SEPMILLIER: THDBValComboBox;
    IM_DELIMITEUR: THDBValComboBox;
    IM_FORMATSENS: THDBValComboBox;
    IM_ANNEE4: TDBCheckBox;
    HDBSpinEdit1: THDBSpinEdit;
    DBCheckBox1: TDBCheckBox;
    FTaille: THDBSpinEdit;
    FAscii: THDBValComboBox;
    DBCheckBox2: TDBCheckBox;
    BVentil: TToolbarButton97;
    TaIM_NATURE: TStringField;
    TaIM_IMPORTATION: TStringField;
    TaIM_CODE: TStringField;
    TaIM_LIBELLE: TStringField;
    TaIM_FORMATDATE: TStringField;
    TaIM_ANNEE4: TStringField;
    TaIM_CRLF: TStringField;
    TaIM_ASCII: TStringField;
    TaIM_SEPDATE: TStringField;
    TaIM_SEPDECIMAL: TStringField;
    TaIM_SEPMILLIER: TStringField;
    TaIM_MOINSDEVANT: TStringField;
    TaIM_DELIMITEUR: TStringField;
    TaIM_IGNORELIGNE: TIntegerField;
    TaIM_TAILLEREC: TIntegerField;
    TaIM_FORMATSENS: TIntegerField;
    IM_PREFIXE: THDBValComboBox;
    HLabel3: THLabel;
    TaIM_PREFIXE: TStringField;
    BCopy: TToolbarButton97;
    TFMTIMPDE: THTable;
    procedure BVentilClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure TaAfterDelete(DataSet: TDataSet);
    procedure TaAfterPost(DataSet: TDataSet);
    procedure TaBeforeDelete(DataSet: TDataSet);
    procedure BDeleteClick(Sender: TObject);
    procedure BCopyClick(Sender: TObject);
  private
    FNat,FImp : String ;
    FAvertir  : boolean ;
    Function  VerifOk : boolean ;
    function  TransChamp(NomChamp : String) : String ;
  public
    { Déclarations publiques }
   // ajout me n'est jamais utilisé Function  CodExist : Boolean ; override ;
    Procedure NewEnreg ; Override ;
  end;

implementation

uses FmtDetai ;

{$R *.DFM}

Function TFFmtImport.VerifOk : boolean ;
BEGIN
Result:=False ;
if (TaIM_CODE.AsString='')     then BEGIN HM2.Execute(4,Caption,'') ; Exit ; END ;
if (TaIM_Prefixe.AsString='')  then BEGIN HM2.Execute(5,Caption,'') ; Exit ; END ;
Result:=True ;
END ;

function TFFmtImport.TransChamp(NomChamp : String) : String ;
var p : Integer ;
    FPrefixe,Pref : String ;
BEGIN
Result:=NomChamp ;
p:=Pos('_',NomChamp) ;
if p<>0 then Pref:=Copy(NomChamp,1,p-1) ;
FPrefixe:=TaIM_Prefixe.AsString ;
//if (Pref<>'IE') or (Pref<>'E') or
if (p=0) or (Pref='XX') then Exit ;
if (Pref='IE') and (FImp='-') then Pref:=FPrefixe ;
if ((FPrefixe='BE') or (FPrefixe='Y') or (FPrefixe='E')) and (FImp='X') then Pref:='IE' ;
Result:=Pref+Trim(Copy(NomChamp,p,length(NomChamp)-p+1)) ;
if (Pref='IE') and (FImp='X') then
  BEGIN
  if NomChamp=FPrefixe+'_NUMEROPIECE' then Result:='IE_NUMPIECE' ;
  if NomChamp='BE_BUDJAL' then Result:='IE_JOURNAL' ;
  if NomChamp='BE_BUDGENE' then Result:='IE_GENERAL' ;
  if NomChamp='BE_BUDSECT' then Result:='IE_SECTION' ;
  if NomChamp='BE_NATUREBUD' then Result:='IE_NATUREPIECE' ;
  END ;
if (Pref=FPrefixe) and (FImp='-') then
  BEGIN
  if NomChamp='IE_NUMPIECE' then Result:=FPrefixe+'_NUMEROPIECE' ;
  if FPrefixe='BE' then
    BEGIN
    if NomChamp='IE_JOURNAL' then Result:=FPrefixe+'_BUDJAL' ;
    if NomChamp='IE_GENERAL' then Result:=FPrefixe+'_BUDGENE' ;
    if NomChamp='IE_SECTION' then Result:=FPrefixe+'_BUDSECT' ;
    if NomChamp='IE_NATUREPIECE' then Result:=FPrefixe+'_NATUREBUD' ;
    END ;
  END ;
END ;

procedure TFFmtImport.BVentilClick(Sender: TObject);
begin
if not VerifOk then Exit ;
If FNat='FCC' Then ParamImportDetailLibre(FNat,FImp,TaIM_Code.AsString,TaIM_Libelle.AsString,TaIM_Prefixe.AsString)
              Else ParamImportDetail(FNat,FImp,TaIM_Code.AsString,TaIM_Libelle.AsString,TaIM_Prefixe.AsString) ;
end;

function QuelTT(Nature,Import : String ; Fixe : boolean) : string ;
BEGIN
Result:='' ;
if Fixe then
   BEGIN
   if Nature='FGE' then BEGIN If Import='X' Then Result:='ttformatGeneraux' Else Result:='ttformatGenerauxExp'; END  else
    if Nature='FSE' then BEGIN If Import='X' Then Result:='ttformatSection' Else Result:='ttformatSectionExp' ; END else
     if Nature='FTI' then BEGIN If Import='X' Then Result:='ttformatTiers' Else Result:='ttformatTiersExp' ; END else
      if (Nature='FEC') then BEGIN If Import='X' Then Result:='ttFormatEcriture' Else Result:='ttFormatEcritureExp' ; END else
       if Nature='FOD' then BEGIN If Import='X' Then Result:='ttFormatEcrAna' Else Result:='ttFormatEcrAnaExp' ; END else
        if Nature='FBE' then BEGIN If Import='X' Then Result:='ttFormatEcrBud' Else Result:='ttFormatEcrBudExp' ; END else
         if Nature='FBA' then BEGIN If Import='X' Then Result:='ttFormatBalance' Else Result:='ttFormatBalanceExp' ; END ;
          if Nature='FCC' then BEGIN If Import='X' Then Result:='' Else Result:='ttFormatEnCoursExp' ; END ;
   Exit ;
   END ;
if Import='-' then
   BEGIN
   if Nature='FGE' then Result:='ttExpParamGen' else
    if Nature='FTI' then Result:='ttExpParamTie' else
     if Nature='FSE' then Result:='ttExpParamSec' else
      if (Nature='FEC') then Result:='ttExpParamEcr' else
       if Nature='FOD' then Result:='ttExpParamOda' else
        if Nature='FBE' then Result:='ttExpParamBud' else
//      'L'      : Result:=ttExpParamBal ;
         if Nature='FBJ' then Result:='ttExpParamJB' else
          if Nature='FBG' then Result:='ttExpParamCB' else
           if Nature='FBS' then Result:='ttExpParamSB' Else
            If Nature='FCC' Then Result:='ttExpParamEcc' ;
   END else
   BEGIN
   if Nature='FGE' then Result:='ttImpParamGen' else
    if Nature='FTI' then Result:='ttImpParamTie' else
     if Nature='FSE' then Result:='ttImpParamSec' else
      if (Nature='FEC') then Result:='ttImpParamEcr' else       // 'I' pour 'IE'
       if Nature='FOD' then Result:='ttImpParamOda' else
        if Nature='FBE' then Result:='ttImpParamBud' else
//      'L'      : Result:=ttExpParamBal ;
         if Nature='FBJ' then Result:='ttImpParamJB' else
          if Nature='FBG' then Result:='ttImpParamCB' else
           if Nature='FBS' then Result:='ttImpParamSB' ;
   END ;
END ;

(* ajout me n'est jamais utilisé
function CodeFormatFixeExiste(FNat : String ; Code : String) : boolean ;
var tt : string ;
    Va,It : TStrings ;
BEGIN
Result:=False ;
tt:=QuelTT(FNat,'-',True) ;
Va:=TStringList.Create ; It:=TStringList.Create ;
RemplirValCombo(tt,'','',It,Va,False,False) ;
Result:=((Va.IndexOf(Code))>-1) ;
Va.Clear ; Va.Free ; It.Clear ; It.Free ;
END ;
*)

Procedure ParamImport (Nature,Import : String ) ;
var FFmtImport: TFFmtImport;
    tt : string ;
begin
//if Blocage(['nrCloture','nrBatch'],True,'nrBatch') then Exit ;
FFmtImport:=TFFmtImport.Create(Application) ;
try
  if Import='-' then FFmtImport.Caption:=FFmtImport.HM2.Mess[1]
                else FFmtImport.Caption:=FFmtImport.HM2.Mess[0] ;
  tt:=QuelTT(Nature,Import,False) ;
  FFmtImport.InitFL('IM','PRT_FMTPARAM','',Nature,taModif,TRUE,FFmtImport.TaIM_CODE,FFmtImport.TaIM_LIBELLE,nil,[tt]) ;
  FFmtImport.FNat:=Nature ;
  FFmtImport.FImp:=Import ;
  FFmtImport.ShowModal ;
  finally
  FFmtImport.Free ;
//  Bloqueur('nrBatch',False) ;
  end ;
SourisNormale ;
end ;
(* ajout me n'est jamais utilisé
Function  TFFmtImport.CodExist : Boolean ;
var Existe : boolean ;
BEGIN
Result:=inherited CodExist ; if Result then Exit ;
if Ta.State in [dsEdit,dsBrowse] then exit ;
Existe:=PresenceComplexe ('FMTIMPOR',['IM_NATURE','IM_IMPORTATION','IM_CODE'],['=','=','='],[FNat,FImp,TaIM_CODE.AsString],['S','S','S']) ;
if Existe then
   BEGIN
   if IM_CODE.CanFocus then IM_CODE.SetFocus ; Result:=True ;
   HM.Execute(4,Caption,'')  ;
   Exit ;
   END ;
if not Existe then Existe:=CodeFormatFixeExiste(FNat,TaIM_CODE.AsString) ;
if Existe then
   BEGIN
   if IM_CODE.CanFocus then IM_CODE.SetFocus ; Result:=True ;
   HM2.Execute(2,Caption,'')  ;
   END ;
Result:= Existe;
END ;
*)

procedure TFFmtImport.FormShow(Sender: TObject);
var i : integer ;
begin
FAvertir:=False ;
IM_PREFIXE.Values.Clear ; IM_PREFIXE.Items.Clear ;
For i:=1 to High(V_PGI.DEtables) do if V_PGI.DETables[i].Prefixe<>'' then
   BEGIN
   IM_PREFIXE.Values.Add(V_PGI.DETables[i].Prefixe) ;
   IM_PREFIXE.Items.Add(V_PGI.DETables[i].Nom) ;
   END ;
  inherited;
if FImp='X' then BCopy.Hint:=HM2.Mess[6]
            else BCopy.Hint:=HM2.Mess[7] ;
Ta.SetRange([FNat,FImp],[FNat,FImp]) ;
if (Ta.Eof) And (Ta.Bof) And (FTypeAction<>taConsult) then
   BEGIN
   if ta.State=dsInsert then NewEnreg else BinsertClick(Nil) ;
   END ;
end;

procedure TFFmtImport.TaAfterDelete(DataSet: TDataSet);
begin
  inherited;
FAvertir:=True ;
end;

procedure TFFmtImport.TaAfterPost(DataSet: TDataSet);
begin
  inherited;
FAvertir:=True ;
end;

Procedure TFFmtImport.NewEnreg ;
BEGIN
Inherited ;
taIM_Nature.AsString:=FNat ;
taIM_IMPORTATION.AsString:=FImp ;
END ;

procedure TFFmtImport.TaBeforeDelete(DataSet: TDataSet);
var Code : String3 ;
begin
Code:=TaIM_CODE.AsString ;
  inherited;
ExecuteSQL('DELETE FROM FMTIMPDE WHERE ID_NATURE="'+FNat+'" AND ID_IMPORTATION="'+FImp+'"'+
           ' AND ID_CODE="'+Code+'"') ;
end;

procedure TFFmtImport.BDeleteClick(Sender: TObject);
begin
if (((TaIM_CODE.AsString='CRA') or (TaIM_CODE.AsString='CVI') or (TaIM_CODE.AsString='CPR')
   or (TaIM_CODE.AsString='CLB') or (TaIM_CODE.AsString='CTR') or (TaIM_CODE.AsString='CRL'))
   and ((TaIM_NATURE.AsString='E') or (TaIM_NATURE.AsString='IE')))
   or (((TaIM_NATURE.AsString='CPR') or (TaIM_NATURE.AsString='CVI') or (TaIM_NATURE.AsString='CLB'))
   and (TaIM_CODE.AsString='T')) then
  BEGIN
  HM2.Execute(3,Caption,'') ;
  Exit ;
  END ;
  inherited;
end;

procedure TFFmtImport.BCopyClick(Sender: TObject);
var Q : TQuery ;
    FIE : String ;
    M,M1 : Integer ;
begin
if not VerifOk then Exit ;
M:=8 ; M1:=10 ;
if FImp='-' then BEGIN Inc(M) ; Inc(M1) ; END ;
// Import dans Impecr / Export dans Ecriture
FIE:='X' ;
if FImp='X' then FIE:='-' ;
// Entete
Q:=OpenSQL('SELECT * FROM FMTIMPOR WHERE IM_NATURE="'+FNat+'" AND IM_IMPORTATION="'+FIE+'" AND IM_CODE="'+TaIM_CODE.AsString+'"',True) ;
if Q.Eof then BEGIN HM2.Execute(M1,Caption,'') ; Ferme(Q) ; Exit ; END ;
if HM2.Execute(M,Caption,'')<>mrYes then BEGIN Ferme(Q) ; Exit ; END ;
if Ta.State=dsBrowse then Ta.Edit ;
TaIM_PREFIXE.AsString:=Q.FindField('IM_PREFIXE').AsString ;
TaIM_LIBELLE.AsString:=Q.FindField('IM_LIBELLE').AsString ;
TaIM_FORMATDATE.AsString:=Q.FindField('IM_FORMATDATE').AsString ;
TaIM_SEPDATE.AsString:=Q.FindField('IM_SEPDATE').AsString ;
TaIM_ANNEE4.AsString:=Q.FindField('IM_ANNEE4').AsString ;
TaIM_ASCII.AsString:=Q.FindField('IM_ASCII').AsString ;
TaIM_SEPMILLIER.AsString:=Q.FindField('IM_SEPMILLIER').AsString ;
TaIM_SEPDECIMAL.AsString:=Q.FindField('IM_SEPDECIMAL').AsString ;
TaIM_DELIMITEUR.AsString:=Q.FindField('IM_DELIMITEUR').AsString ;
TaIM_IGNORELIGNE.AsString:=Q.FindField('IM_IGNORELIGNE').AsString ;
TaIM_FORMATSENS.AsString:=Q.FindField('IM_FORMATSENS').AsString ;
TaIM_MOINSDEVANT.AsString:=Q.FindField('IM_MOINSDEVANT').AsString ;
TaIM_CRLF.AsString:=Q.FindField('IM_CRLF').AsString ;
TaIM_TAILLEREC.AsString:=Q.FindField('IM_TAILLEREC').AsString ;
Ferme(Q) ;
// détail du format
Q:=OpenSQL('SELECT * FROM FMTIMPDE WHERE ID_NATURE="'+FNat+'" AND ID_IMPORTATION="'+FIE+'" AND ID_CODE="'+TaIM_CODE.AsString+'"'+
           ' ORDER BY ID_NATURE, ID_IMPORTATION, ID_CODE, ID_NUMLIGNE, ID_NUMCHAMP',TRUE) ;
ExecuteSQL('DELETE FROM FMTIMPDE WHERE ID_NATURE="'+FNat+'" AND ID_IMPORTATION="'+FImp+'" AND ID_CODE="'+TaIM_CODE.AsString+'"') ;
TFMTIMPDE.Open ;
While Not Q.EOF do
  BEGIN
  TFMTIMPDE.Insert ;
  TFMTIMPDE.FindField('ID_NATURE').AsString:=FNat ;
  TFMTIMPDE.FindField('ID_IMPORTATION').AsString:=FImp ;
  TFMTIMPDE.FindField('ID_CODE').AsString:=Q.FindField('ID_CODE').AsString ;
  TFMTIMPDE.FindField('ID_NUMLIGNE').AsInteger:=Q.FindField('ID_NUMLIGNE').AsInteger ;
  TFMTIMPDE.FindField('ID_NUMCHAMP').AsInteger:=Q.FindField('ID_NUMCHAMP').AsInteger ;
  TFMTIMPDE.FindField('ID_CHAMP').AsString:=TransChamp(Q.FindField('ID_CHAMP').AsString) ;
  TFMTIMPDE.FindField('ID_DEBUT').AsInteger:=Q.FindField('ID_DEBUT').AsInteger ;
  TFMTIMPDE.FindField('ID_LONGUEUR').AsInteger:=Q.FindField('ID_LONGUEUR').AsInteger ;
  TFMTIMPDE.FindField('ID_DECIMAL').AsInteger:=Q.FindField('ID_DECIMAL').AsInteger ;
  TFMTIMPDE.FindField('ID_CORRESP').AsString:=Q.FindField('ID_CORRESP').AsString ;
  TFMTIMPDE.FindField('ID_ALIGNEDROITE').AsString:=Q.FindField('ID_ALIGNEDROITE').AsString ;
  TFMTIMPDE.Post ; Q.Next ;
  END ;
TFMTIMPDE.Close ;
Ferme(Q) ;
BValiderClick(nil) ;
end;

end.
