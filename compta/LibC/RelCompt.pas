unit RelCompt;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, hmsgbox, Hctrls, StdCtrls, Mask, Menus, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  Hqry, Grids,
  DBGrids, ExtCtrls, ComCtrls, Buttons, HEnt1, Ent1, HStatus,
  CPTiers_TOM,
{$IFDEF V530}
     EdtDoc,
{$ELSE}
     EdtRDoc,EdtDoc,
{$ENDIF}
  ZEcriMvt_TOF, Filtre, HRichEdt, HSysMenu, Hcompte, HDB,  LettUtil, HTB97,
  ColMemo, HPanel, UiUtil, SaisUtil, HRichOLE, ed_Tools, ADODB ;

procedure ReleveCompte ;

type
  TFRelCompt = class(TFMul)
    H_AUXILIAIRE: THLabel;
    T_AUXILIAIRE_SUP: THCpteEdit;
    H_AUXILIAIRE_: THLabel;
    T_AUXILIAIRE_INF: THCpteEdit;
    XX_WHERE: TEdit;
    T_LIBELLE: TEdit;
    HLabel1: THLabel;
    TE_DATECOMPTABLE: THLabel;
    TE_DATECOMPTABLE2: THLabel;
    APP: TCheckBox;
    ORDREEDIT: THValComboBox;
    H_OrdreEdit: THLabel;
    BMenuZoom: TToolbarButton97;
    HM: THMsgBox;
    BZoomPiece: THBitBtn;
    POPZ: TPopupMenu;
    ChoixNat: TRadioGroup;
    HLabel3: THLabel;
    PLibres: TTabSheet;
    Bevel5: TBevel;
    TT_TABLE0: THLabel;
    TT_TABLE1: THLabel;
    TT_TABLE2: THLabel;
    TT_TABLE3: THLabel;
    TT_TABLE4: THLabel;
    TT_TABLE5: THLabel;
    TT_TABLE6: THLabel;
    TT_TABLE7: THLabel;
    TT_TABLE8: THLabel;
    TT_TABLE9: THLabel;
    T_TABLE4: THCpteEdit;
    T_TABLE3: THCpteEdit;
    T_TABLE2: THCpteEdit;
    T_TABLE1: THCpteEdit;
    T_TABLE0: THCpteEdit;
    T_TABLE5: THCpteEdit;
    T_TABLE6: THCpteEdit;
    T_TABLE7: THCpteEdit;
    T_TABLE8: THCpteEdit;
    T_TABLE9: THCpteEdit;
    BZoom: THBitBtn;
    HLabel9: THLabel;
    _DEVISE: THValComboBox;
    H_SECTEUR: THLabel;
    MODELE: THValComboBox;
    BParamModele: TToolbarButton97;
    HLabel2: THLabel;
    _ETABLISSEMENT: THMultiValComboBox;
    TE_EXERCICE: THLabel;
    _EXERCICE: THValComboBox;
    AvecLettrage: TCheckBox;
    _DATECOMPTABLE: THCritMaskEdit;
    _DATECOMPTABLE_: THCritMaskEdit;
    HLabel4: THLabel;
    HLabel5: THLabel;
    H_COLLECTIF: THLabel;
    H_COLLECTIF_: THLabel;
    T_COLLECTIF_SUP: THCpteEdit;
    T_COLLECTIF_INF: THCpteEdit;
    procedure FormShow(Sender: TObject);
    procedure BZoomClick(Sender: TObject);
    procedure BParamModeleClick(Sender: TObject);
    procedure BZoomPieceClick(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject); override;
    procedure FListeDblClick(Sender: TObject); override;
    procedure ChoixNatClick(Sender: TObject);
    procedure BChercheClick(Sender: TObject); override;
    procedure FormCreate(Sender: TObject); override;
    procedure _EXERCICEChange(Sender: TObject);
    procedure BMenuZoomMouseEnter(Sender: TObject);
  private    { Déclarations privées }
    function  JeValide : boolean ;
    function  FabricSQLEcr : String ;
    function  OkMvt ( SQL : String ) : boolean ;
    procedure EditeRelevesCompte ;
    function  OrdreTri : String ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcGen,
  ULibExercice,
  {$ENDIF MODENT1}
  UtilEdt, UtilPgi,URecupSQLModele  ;

procedure ReleveCompte ;
Var X  : TFRelCompt ;
    PP : THPanel ;
BEGIN
if _Blocage(['nrCloture'],False,'nrAucun') then Exit ;
PP:=FindInsidePanel ;
X:=TFRelCompt.Create(Application) ;
X.FNomFiltre:='RELEVECOMPTE' ; X.Q.Manuel:=True ; X.Q.Liste:='RELEVECOMPTE' ;
if PP=Nil then
   BEGIN
    try
     X.ShowModal ;
    Finally
     X.Free ;
    End ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
END ;

procedure TFRelCompt.FormShow(Sender: TObject);
//Var st,StV8 : String ;
begin

    HelpContext:=7589000;

if VH^.CPExoRef.Code<>'' then
   BEGIN
   _EXERCICE.Value:=VH^.CPExoRef.Code ;
   _DATECOMPTABLE.Text:=DateToStr(VH^.CPExoRef.Deb) ;
   _DATECOMPTABLE_.Text:=DateToStr(VH^.CPExoRef.Fin) ;
   END else
   BEGIN
   _EXERCICE.Value:=VH^.Entree.Code ;
   _DATECOMPTABLE.Text:=StDate1900 ;
   _DATECOMPTABLE_.Text:=StDate2099 ;
   END ;
InitTablesLibresTiers(PLibres) ;
_DEVISE.Value:=V_PGI.DevisePivot ; ORDREEDIT.ItemIndex:=0 ;
if MODELE.Values.Count>0 then MODELE.Value:=MODELE.Values[0] ;
  inherited;
Q.Manuel:=False ; BChercheClick(nil);
//RR Pages.Pages[3].TabVisible:=FALSE ;
//Pages.Pages[4].TabVisible:=FALSE ;
PositionneEtabUser(_ETABLISSEMENT, False); // 15088
end;

procedure TFRelCompt.BZoomClick(Sender: TObject);
Var Cpte : String ;
begin
  inherited;
if ((Q.EOF) and (Q.BOF)) then Exit ; Cpte:=Q.FindField('T_AUXILIAIRE').AsString ;
FicheTiers(Q,'',Cpte,taConsult,1) ;
end;

procedure TFRelCompt.BParamModeleClick(Sender: TObject);
begin
  inherited;
EditDocumentS5S7('L','RLC',MODELE.Value,True) ;
end;

procedure TFRelCompt.BZoomPieceClick(Sender: TObject);
Var Auxi : String ;
begin
  inherited;
if ((Q.EOF) and (Q.BOF)) then Exit ; Auxi:=Q.FindField('T_AUXILIAIRE').AsString ;
ZoomEcritureMvt(Auxi,fbAux,'MULMMVTS' ) ;
end;

function TFRelCompt.JeValide : boolean ;
Var ii : integer ;
BEGIN
Result:=False ;
if ((Q.EOF) and (Q.BOF)) then BEGIN HM.Execute(0,'','') ; Exit ; END ;
if FListe.NbSelected<=0 then BEGIN HM.Execute(0,'','') ; Exit ; END ;
if MODELE.Value='' then BEGIN HM.Execute(1,'','') ; Exit ; END ;
ii:=HM.Execute(2,'','') ;
Case ii of
   mrNo     : BEGIN FListe.ClearSelected ; Exit ; END ;
   mrCancel : Exit ;
   END ;
Result:=True ;
END ;

function TFRelCompt.FabricSQLEcr : String ;
Var St,St2,StOr,St3,StV8 : String ;
BEGIN
St := '';
if T_COLLECTIF_SUP.text<>'' then St := 'E_GENERAL>="' + T_COLLECTIF_SUP.text + '" AND ';
if T_COLLECTIF_INF.text<>'' then St := St + 'E_GENERAL<="' + T_COLLECTIF_INF.text + '" AND ';
St:=St+' E_DATECOMPTABLE>="'+UsDate(_DATECOMPTABLE)+'" AND E_DATECOMPTABLE<="'+UsDate(_DATECOMPTABLE_)+'"' ;
St:=St+' AND E_DEVISE="'+_DEVISE.Value+'" AND E_QUALIFPIECE="N" AND E_TRESOLETTRE<>"X" ' ;
if _EXERCICE.Value<>'' then St:=St+' AND E_EXERCICE="'+_EXERCICE.Value+'"' ;
St:=St+' AND ((E_FLAGECR<>"ROR" AND E_FLAGECR<>"RDE") OR (E_FLAGECR="")) ' ;
if (_ETABLISSEMENT.Text<>'') And (_ETABLISSEMENT.Text<>Traduirememoire('<<Tous>>')) then
   BEGIN
   St2:=_ETABLISSEMENT.Text ; StOr:='' ;
   While St2<>'' do StOr:=StOr+' OR E_ETABLISSEMENT="'+ReadTokenSt(St2)+'"' ;
   System.Delete(StOr,1,4) ;
   if StOR<>'' then St:=St+' AND ('+StOr+')' ;
   END ;
If Not AvecLettrage.Checked Then St:=St+' AND E_ETATLETTRAGE<>"TL" ' ;
St3:='E_ECRANOUVEAU="N"' ;
StV8:=LWhereV8 ;
if StV8<>'' then
  BEGIN
  St3:='('+St3+' OR E_ECRANOUVEAU="H")  ' ;
  St3:=St3+' AND ('+StV8+') ' ;
  END ;
St:=St+' AND '+St3 ;
Result:=St ;
END ;

function TFRelCompt.OkMvt ( SQL : String ) : boolean ;
Var QQ : TQuery ;
BEGIN
QQ:=OpenSQL(SQL,True) ;
Result:=(Not QQ.EOF) ;
Ferme(QQ) ; 
END ;

procedure TFRelCompt.EditeRelevesCompte;
var
    i,j : integer;
    WhereEcr,Sql,StOrder : string;
    Auxi,Model : string;
    TT : TList;
    LL : TStrings;
begin
    WhereEcr := FabricSQLEcr;
    TT := Tlist.Create;

    InitMove(FListe.NbSelected,'');
    for i := 0 to FListe.NbSelected-1 do
    begin
        MoveCur(FALSE);
        FListe.GotoLeBookMark(i);

        Auxi := Q.FindField('T_AUXILIAIRE').AsString;
        Model := Modele.Value;
        LL:=TStringList.Create;
        Sql:=RecupSQLModele('L','RLC',Model,'',Auxi,'');
        j:=Pos('ORDER BY',Sql);
        if (j > 0) then
        begin
            StOrder := Copy(Sql,j,Length(Sql)-j+1);
            System.Delete(Sql,j,Length(Sql)-j+1);
        end
        else StOrder := '';

        if (Trim(WhereEcr) <> '') then Sql := Sql + ' AND ' + WhereEcr;
        if (StOrder <> '') then Sql := Sql + ' ' + StOrder;
        if (OkMvt(SQL)) then
        begin
            LL.Add(Sql);
            TT.Add(LL);
        end
        else LL.Free;
    end;

    FiniMove;
    //LanceDoc('L','RLC',MODELE.Value,TT,Nil,APP.Checked,True);
    LanceDocument('L','RLC',MODELE.Value,TT,Nil,APP.Checked,True);
    VideListe(TT);
    TT.Free;
end;

procedure TFRelCompt.BOuvrirClick(Sender: TObject);
begin
//  inherited;
if Not JeValide then Exit ;
Application.ProcessMessages ;
EditeRelevesCompte ; FListe.ClearSelected ;
end;

procedure TFRelCompt.FListeDblClick(Sender: TObject);
begin
  inherited;
BZoomPieceClick(Nil) ;
end;

procedure TFRelCompt.ChoixNatClick(Sender: TObject);
begin
  inherited;
if ChoixNat.ItemIndex=0 then XX_WHERE.Text:='T_NATUREAUXI="CLI" or T_NATUREAUXI="AUD"'
                        else XX_WHERE.Text:='T_NATUREAUXI="FOU" or T_NATUREAUXI="AUC"' ;
{ FQ 15644 - CA - 14/06/2005 }
  FListe.ClearSelected; // Si on change de nature ==> on efface la sélection
end;

procedure TFRelCompt.BChercheClick(Sender: TObject);
var St : String ;
    i  : integer ;
begin
Q.DisableControls ;
St:='' ; Q.Close ;
Q.UpdateCriteres ;
for i:=0 to Q.SQL.Count-1 do St:=St+Q.SQL[i] ;
i:=Pos('ORDER BY',St) ;
if i<>0 then System.Delete(St,i,Length(St)-i+1) ;
St:=St+' '+OrdreTri ;
Q.SQL.Clear ; Q.SQL.Add(St) ;
Q.Open ; {attention ChangeSQL déjà fait par UpdateCriteres}
Q.EnableControls ;
end;

function TFRelCompt.OrdreTri : String ;
BEGIN                                          
Result:='' ;
if ORDREEDIT.Value='CPT' then
   BEGIN
   Result:='ORDER BY T_AUXILIAIRE,T_LIBELLE' ;
   END else if ORDREEDIT.Value='INT' then
   BEGIN
   Result:='ORDER BY T_LIBELLE,T_AUXILIAIRE' ;
   END else if ORDREEDIT.Value='PCP' then
   BEGIN
   Result:='ORDER BY T_PAYS, T_CODEPOSTAL, T_AUXILIAIRE' ;
   END ;
END ;

procedure TFRelCompt.FormCreate(Sender: TObject);
begin
  inherited;
{$IFDEF CCMP}
if (VH^.CCMP.LotCli) then begin ChoixNat.ItemIndex:=0; ChoixNat.Visible:=False; HLabel3.Visible:=False; end
                     else begin ChoixNat.ItemIndex:=1; ChoixNat.Visible:=False; HLabel3.Visible:=False; end;
{$ENDIF}
end;

procedure TFRelCompt._EXERCICEChange(Sender: TObject);
begin
  inherited;
  ExoToDates(_EXERCICE.Value,_DATECOMPTABLE,_DATECOMPTABLE_);
  if _EXERCICE.Value='' then
  begin                                           
    _DATECOMPTABLE.Text:=stDate1900;
    _DATECOMPTABLE_.Text:=stDate2099;
  end;

end;

procedure TFRelCompt.BMenuZoomMouseEnter(Sender: TObject);
begin
  inherited;
PopZoom97(BMenuZoom,POPZ) ;
end;

end.
