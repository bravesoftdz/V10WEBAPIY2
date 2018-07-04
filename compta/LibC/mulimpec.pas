unit mulimpec;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, hmsgbox, Hcompte, Mask, Hctrls, StdCtrls, HSysMenu, Menus, Db,
  DBTables, Hqry, HTB97, ComCtrls, HRichEdt, ExtCtrls, Grids, DBGrids, HDB,
  ColMemo, hent1, Ent1, HRichOLE;

procedure MultiCritereImpEcr(Comment : TActionFiche ; ListeDesErreurs : TList);

type
  TFMulMvtImp = class(TFMul)
    TIE_JOURNAL: THLabel;
    TIE_NATUREPIECE: THLabel;
    IE_NATUREPIECE: THValComboBox;
    IE_JOURNAL: THValComboBox;
    TIE_NUMEROPIECE: THLabel;
    IE_NUMPIECE: THCritMaskEdit;
    HLabel1: THLabel;
    IE_NUMPIECE_: THCritMaskEdit;
    TIE_DATECOMPTABLE: THLabel;
    IE_DATECOMPTABLE: THCritMaskEdit;
    TE_DATECOMPTABLE2: THLabel;
    IE_DATECOMPTABLE_: THCritMaskEdit;
    TIE_ETABLISSEMENT: THLabel;
    IE_ETABLISSEMENT: THValComboBox;
    TIE_GENERAL: THLabel;
    IE_GENERAL: THCpteEdit;
    TIE_AUXILIAIRE: THLabel;
    IE_AUXILIAIRE: THCpteEdit;
    HM: THMsgBox;
    RUne: TCheckBox;
    XX_WHERE: TEdit;
    IE_NUMLIGNE_: THCritMaskEdit;
    IE_NUMLIGNE: THCritMaskEdit;
    IE_NUMECHE: THCritMaskEdit;
    HCpte: THCpteEdit;
    TIE_REFINTERNE: THLabel;
    IE_REFINTERNE: TEdit;
    TIE_SECTION: THLabel;
    IE_SECTION: THCpteEdit;
    RMVTANA: TCheckBox;
    RMVTGEN: TCheckBox;
    IE_ANA: THCritMaskEdit;
    TIE_DEVISE: THLabel;
    IE_DEVISE: THValComboBox;
    TIE_AXE: THLabel;
    IE_AXE: THValComboBox;
    EBListeErreur: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure RUneClick(Sender: TObject);
    procedure FListeColEnter(Sender: TObject);
    procedure FListeDblClick(Sender: TObject); override;
    procedure RMVTANAClick(Sender: TObject);
    procedure RMVTGENClick(Sender: TObject);
    procedure EBListeErreurClick(Sender: TObject);
  private
//    ttListe : TList ;
    ListeDesErreurs : TList ;
//    procedure ChargeListeTT(Nom : String ; ttCorr : String) ;
//    Procedure ChargePickList(i : byte ; ttCorr : String) ;
  public

  end;

implementation

{$R *.DFM}

Uses ImpUtil,RapSuppr ;

type TFttListe = Class
     TypeTable : String ;
     Code,Lib    : String ;
     END ;


procedure MultiCritereImpEcr(Comment : TActionFiche ; ListeDesErreurs : TList);
var
  FMulMvt : TFMulMvtImp ;
begin
FMulMvt:=TFMulMvtImp.Create(Application) ;
try
  FMulMvt.TypeAction:=Comment ;
  FMulMvt.FNomFiltre:='INTMVT' ;
  FMulMvt.ListeDesErreurs:=ListeDesErreurs ;
  FMulMvt.Q.Manuel:=TRUE ;
  Case Comment Of
    taConsult : begin
                FMulMvt.Caption:=FMulMvt.HM.Mess[0] ;
                end ;
    taModif   : begin
                FMulMvt.Caption:=FMulMvt.HM.Mess[1] ;
                end ;
    end ;
  FMulMvt.Q.Liste:='INTMVT' ;
  FMulMvt.ShowModal ;
  finally
  FMulMvt.Free ;
  end;
Screen.Cursor:=SyncrDefault ;
end;

function Max(X, Y: Integer): Integer;
begin
  Result := Y;
  if X > Y then Result := X;
end;


procedure TFMulMvtImp.FormShow(Sender: TObject);
begin
IE_DATECOMPTABLE.Text:=DateToStr(VH^.Encours.Deb) ;
If VH^.Suivant.Code<>'' Then IE_DATECOMPTABLE_.Text:=DateToStr(VH^.Suivant.Fin)
                        Else IE_DATECOMPTABLE_.Text:=DateToStr(VH^.EnCours.Fin)  ;
  inherited;
Q.Manuel:=FALSE ; Q.UpdateCriteres ;
CentreDBGrid(FListe) ;
EBListeErreur.Visible:=ListeDesErreurs.Count>0 ;
end;

procedure TFMulMvtImp.RUneClick(Sender: TObject);
begin
  inherited;
If RUne.Checked Then BEGIN IE_NUMECHE.text:='1' ; IE_NUMLIGNE.text:='0' ; IE_NUMLIGNE_.text:='1' ; END
                Else BEGIN IE_NUMECHE.text:='9999' ; IE_NUMLIGNE.text:='0' ; IE_NUMLIGNE_.text:='9999' ; END ;
end;

// tt,tz des zones du grid
{Procedure TFMulMvtImp.ChargePickList(i : byte ; ttCorr : String) ;
var j : integer ;
    FttListe   : TFttListe ;
BEGIN
exit ;
With FListe.Columns.Items[i] do
  BEGIN
   PickList.Clear ; Picklist.Add('');
   for j:=0 to ttListe.Count-1 do
      BEGIN
      FttListe:=ttListe[j] ;
      if (FttListe.TypeTable=ttCorr) then Picklist.Add(FttListe.Code);
      END ;
  END ;
END ;}

{procedure TFMulMvtImp.ChargeListeTT(Nom : String ; ttCorr : String) ;
var QQ : TQuery ;
    FttListe : TFttListe ;
    OkCCO : boolean ;
    i,P : integer ;
    StTable,StPrefixe,StCode,StWhere,StOrder,SQL,StLib : String ;
BEGIN
exit ;
if ttListe=nil then ttListe:=TList.Create ;
for i:=0 to ttListe.Count-1 do
  BEGIN
  FttListe:=ttListe[i] ; if FttListe.TypeTable=ttCorr then Exit ;
  END ;
GetCorrespType (StTable,ttCorr,StCode,StWhere,StPrefixe,StLib) ;
OkCCO:=((StTable='COMMUN') or (StTable='CHOIXCOD')) ;
if OKCCO then SQL:='Select '+StCode+', '+StLib+', '+StPrefixe+'_ABREGE FROM '+StTable
         else SQL:='Select '+StCode+', '+StLib+' FROM '+StTable ;
if ((StWhere<>'') AND (Pos('="'+W_W+'"',StWhere)>0)) then
   BEGIN
   StWhere:=FindEtReplace(StWhere,'="'+W_W+'"','<>"'+W_W+'"',TRUE) ;
   END ;
if StWhere<>'' then SQL:=SQL+' Where '+StWhere ;
P:=Pos('DISTINCT',StCode) ; if P>0 then StOrder:=Copy(StCode,p+9,50) else StOrder:=StCode ;
if OKCCO then SQL:=SQL+' ORDER BY '+StPrefixe+'_TYPE, '+StOrder
        else SQL:=SQL+' ORDER BY '+StOrder ;
QQ:=OpenSQL(SQL,TRUE) ;
While not QQ.EOF do
  BEGIN
  FttListe:=TFttListe.Create ; FttListe.TypeTable:=ttCorr ;
  FttListe.Code:=QQ.Fields[0].AsString ; FttListe.Lib:=QQ.Fields[1].AsString ;
  ttListe.Add(FttListe) ;
  QQ.Next ;
  END ;
Ferme(QQ) ;
END ;}

procedure TFMulMvtImp.FListeColEnter(Sender: TObject);
var ttCorr : String ;
    Nom : String ;
    i : integer ;
    Masque : String ;
    fb : TFichierBase ;
begin
  inherited;
Nom:=FListe.SelectedField.FieldName ; ttCorr:=Get_Join(Nom) ;
if (ttCorr<>'') and (ttCorr<>'TTJOURNAL') then
  BEGIN
  (*
  ChargeListeTT(Nom,ttCorr) ;
  ChargePickList(FListe.SelectedIndex,ttCorr) ;
  *)
  Q.FindField(Nom).EditMask:='>aaa;0; ' ;
  END ;
Masque:='>' ;
if Nom='IE_SECTION' then
  BEGIN
  if (Q.FindField('IE_AXE')=nil) then Exit ;
  fb:=fbAxe1 ;
  if Q.FindField('IE_AXE').AsString='A1' then Fb:=FbAxe1 else
  if Q.FindField('IE_AXE').AsString='A2' then Fb:=FbAxe2 else
  if Q.FindField('IE_AXE').AsString='A3' then Fb:=FbAxe3 else
  if Q.FindField('IE_AXE').AsString='A4' then Fb:=FbAxe4 else
  if Q.FindField('IE_AXE').AsString='A5' then Fb:=FbAxe5 ;
  for i:=1 to VH^.Cpta[fb].Lg do Masque:=Masque+'a' ;
  Q.FindField(Nom).EditMask:=Masque ;
  END else
  BEGIN
  if Nom='IE_JOURNAL'    then BEGIN for i:=1 to 3 do Masque:=Masque+'a' ; END else
  if Nom='IE_GENERAL'    then BEGIN for i:=1 to VH^.Cpta[fbGene].Lg do Masque:=Masque+'a' ; END else
  if Nom='IE_AUXILIAIRE' then BEGIN for i:=1 to VH^.Cpta[fbAux].Lg do Masque:=Masque+'a' ; END else
//if Nom='IE_BUDGET'  then BEGIN for i:=0 to V_PGI.Cpta[fbBudget].Lg do Masque:=Masque+'a' ; END else
//  if Q.FindField(Nom).DataType=ftFloat then BEGIN Masque:='# ###.####' ; END else
  if (Q.FindField(Nom).DataType=ftDateTime) or (Q.FindField(Nom).DataType=ftDate) then BEGIN Masque:='!99/99/0000;1;_' ; END else
  if Q.FindField(Nom).DataType=ftString then BEGIN for i:=1 to Q.FindField(Nom).DataSize-1 do Masque:=Masque+'c' ; END ; // else Masque:=Masque+';0; ' ;
  if Masque<>'>' then Q.FindField(Nom).EditMask:=Masque ;
  END ;
end;

procedure TFMulMvtImp.FListeDblClick(Sender: TObject);
var RefreshListe,Ok : Boolean ;
    pp :TProcZoom ;
    Nom :String ;
begin
  inherited;
Ok:=True ; pp:=nil ;
Nom:=FListe.SelectedField.FieldName ;
if (Nom='IE_JOURNAL') then BEGIN HCpte.ZoomTable:=TzJournal ; pp:=ProcZoomJal ; HCpte.Bourre:=FALSE ; END else
if (Nom='IE_GENERAL') then BEGIN HCpte.ZoomTable:=TzGeneral ; pp:=ProcZoomGene ; HCpte.Bourre:=TRUE ; END else
if (Nom='IE_AUXILIAIRE') then BEGIN HCpte.ZoomTable:=TzTiers ; pp:=ProcZoomTiers ; HCpte.Bourre:=TRUE ; END else
if (Nom='IE_SECTION') then
   BEGIN
   if (Q.FindField('IE_AXE')=nil) then BEGIN HM.Execute(21,'','') ; Exit ; END ;
   if (Q.FindField('IE_AXE').AsString<>'')
      then HCpte.ZoomTable:=AxeToTz(Q.FindField('IE_AXE').AsString)
      else BEGIN HM.Execute(13,'','') ; Exit ; END ;
   pp:=ProcZoomSection ; HCpte.Bourre:=TRUE ;
   END else
if (Nom='IE_BUDGET') then BEGIN HCpte.ZoomTable:=TzBudgen ; pp:=ProcZoomBudgen ; HCpte.Bourre:=FALSE ; END else Ok:=False ;
if Ok then
  BEGIN
  FListe.SelectedIndex:=FListe.SelectedIndex ;
  HCpte.Text:=Trim(FListe.SelectedField.AsString) ;
  if GChercheCompte(HCpte,pp) then
    BEGIN
    Q.Edit ; FListe.Fields[FListe.SelectedIndex].AsString:=HCpte.Text ; Q.Post ;
    RefreshListe:=False ;
    case HCpte.ZoomTable of
      TzGeneral : (*RefreshListe:=MajLettrageEtVentileGen(HCpte.Text)*) ;
      TzTiers   : (*RefreshListe:=MajEtatLettrageTiers(HCpte.Text)*) ;
      END ;
    if RefreshListe then BEGIN Q.Edit ; Q.Post ; END ;
    END ;
  END;
end;

procedure TFMulMvtImp.RMVTANAClick(Sender: TObject);
begin
  inherited;
If RMVTANA.Checked Then BEGIN IE_ANA.text:='"X"' ; END
                   Else BEGIN IE_ANA.Text:='' ; END ;
end;

procedure TFMulMvtImp.RMVTGENClick(Sender: TObject);
begin
  inherited;
If RMVTGEN.Checked Then BEGIN IE_ANA.text:='"-"' ; END
                   Else BEGIN IE_ANA.Text:='' ; END ;
end;

procedure TFMulMvtImp.EBListeErreurClick(Sender: TObject);
Var GGRien : Boolean ;
begin
  inherited;
GGrien:=FALSE ;
If ListeDesErreurs.Count>0 Then
   RapportdErreurMvt(ListeDesErreurs,3,GGRien,FALSE) ;
END ;

end.
