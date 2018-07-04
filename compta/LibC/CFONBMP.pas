unit CFONBMP ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, HSysMenu, Menus, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  Hqry, StdCtrls, Grids, DBGrids, HDB,
  ComCtrls, HRichEdt, Hctrls, ExtCtrls, Buttons, Mask, Hcompte, Ent1, HEnt1,
  ParamDat, SaisUtil, SaisComm, hmsgbox, CFONB, Saisie, TofVerifRib,
{$IFDEF V530}
      EdtEtat,
{$ELSE}
      EdtREtat,
{$ENDIF}
  LettUtil,
  HTB97, ed_tools, ColMemo, HPanel, UiUtil,
{$IFDEF V530}
     EdtDoc,
{$ELSE}
     EdtRDoc,
{$ENDIF}
  HRichOLE, ADODB ;

procedure ExportCFONBMP ( smp : TSuiviMP ) ;

type
  TFCFONBMP = class(TFMul)
    E_DEVISE: THValComboBox;
    TE_EXERCICE: THLabel;
    E_EXERCICE: THValComboBox;
    E_DATECOMPTABLE_: THCritMaskEdit;
    HLabel6: THLabel;
    E_DATECOMPTABLE: THCritMaskEdit;
    HLabel3: THLabel;
    cExport: TCheckBox;
    cRIB: TCheckBox;
    E_ECHE: TEdit;
    E_QUALIFPIECE: TEdit;
    E_NUMECHE: THCritMaskEdit;
    XX_WHERE1: TEdit;
    E_AUXILIAIRE: THCritMaskEdit;
    E_AUXILIAIRE_: THCritMaskEdit;
    bRib: TToolbarButton97;
    XX_WHERERIB: TEdit;
    XX_WHEREEXPORT: TEdit;
    XX_WHEREAUX: TEdit;
    HM: THMsgBox;
    E_ECRANOUVEAU: TEdit;
    XX_WHERENAT: TEdit;
    XX_WHEREDC: TEdit;
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
    TE_JOURNAL: THLabel;
    E_JOURNAL: THValComboBox;
    c: THLabel;
    E_MODEPAIE: THMultiValComboBox;
    TE_NATUREPIECE: THLabel;
    E_NATUREPIECE: THValComboBox;
    HAuxiliaire1: THLabel;
    Auxiliaire1: THCpteEdit;
    HAuxiliaire2: THLabel;
    Auxiliaire2: THCpteEdit;
    HLabel10: THLabel;
    E_DATEECHEANCE: THCritMaskEdit;
    HLabel7: THLabel;
    E_DATEECHEANCE_: THCritMaskEdit;
    HLabel2: THLabel;
    E_GENERAL: THCpteEdit;
    TE_NUMEROPIECE: THLabel;
    E_NUMEROPIECE: THCritMaskEdit;
    HLabel4: THLabel;
    E_NUMEROPIECE_: THCritMaskEdit;
    FExport: TCheckBox;
    TE_DEVISE: THLabel;
    CTIDTIC: TCheckBox;
    XX_WHERENATTIERS: TEdit;
    BCtrlRib: TToolbarButton97;                     // BPY 29/01/2003 Correction de la fiche 10622
    procedure FormShow(Sender: TObject);
    procedure Auxiliaire1Change(Sender: TObject);
    procedure Auxiliaire2Change(Sender: TObject);
    procedure E_EXERCICEChange(Sender: TObject);
    procedure E_DATECOMPTABLEKeyPress(Sender: TObject; var Key: Char);
    procedure bRibClick(Sender: TObject);
    procedure SQDataChange(Sender: TObject; Field: TField);
    procedure cRIBClick(Sender: TObject);
    procedure cExportClick(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject); override;
    procedure FListeDblClick(Sender: TObject); override;
    procedure CTIDTICClick(Sender: TObject);
    procedure BCtrlRibClick(Sender: TObject);
  private
    BanqueGene : String ;
    Function  GetLeOBM : TOBM ;
    Function  ExporteSelection : boolean ;
//    procedure ChargeWhereTraite ;
//    procedure CompleteTL ( TL : Tlist ) ;
//    Procedure FlagExportLettre ( TL : Tlist ) ;
    Function  CoherBanque ( TL : TList ) : boolean ;
    Procedure InitTitres ;
  public
    smp : TSuiviMP ;
    isCli : boolean ;
  end;


implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  ULibExercice,
  {$ENDIF MODENT1}
  DocRegl;

Function SMPClient ( smp : TSuiviMP ) : boolean ;
BEGIN
Result:=(smp in [smpEncTraEnc,smpEncTraEsc,smpEncPreBqe]) ;
END ;


procedure ExportCFONBMP ( smp : TSuiviMP ) ;
Var X  : TFCFONBMP ;
    PP : THPanel ;
BEGIN
if Blocage(['nrCloture','nrBatch','nrLettrage'],True,'nrAucun') then Exit ;
PP:=FindInsidePanel ;
X:=TFCFONBMP.Create(Application) ;
X.XX_WHERE1.Text:='E_JOURNAL="zzz"' ;
X.smp:=smp ; X.isCli:=smpClient(smp) ;
if X.isCli then
   BEGIN
   X.Q.Liste:='CPCFONBMPCLI' ;
   X.FNomFiltre:='MULCFONBMPCLI' ;
   END else
   BEGIN
     if (smp in [smpDecVirInBqe]) then begin
       X.Q.Liste:='CPCFONBINMPFOU' ;
       X.FNomFiltre:='MULCFONBINMPFOU' ;
       X.BCtrlRib.visible := FALSE;         // BPY 29/01/2003 Correction de la fiche 10622
       end
     else begin
       X.Q.Liste:='CPCFONBMPFOU' ;
       X.FNomFiltre:='MULCFONBMPFOU' ;
     end;
   END ;
if PP=Nil then
   BEGIN
    try
     X.ShowModal ;
    finally
     X.Free ;
    end;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
END ;

{Procedure TFCFONBMP.ChargeWhereTraite ;
Var QQ : TQuery ;
    St,SQL : String ;
BEGIN
Case smp of
   smpEnctraEnc,smpEncTraEsc : SQL:='Select MP_MODEPAIE FROM MODEPAIE WHERE MP_CATEGORIE="LCR"' ;
                smpEncPreBqe : SQL:='Select MP_MODEPAIE FROM MODEPAIE WHERE MP_CATEGORIE="PRE"' ;
                smpDecBorDec : SQL:='Select MP_MODEPAIE FROM MODEPAIE WHERE MP_CATEGORIE="LCR"' ;
                smpDecVirbqe : SQL:='Select MP_MODEPAIE FROM MODEPAIE WHERE MP_CATEGORIE="VIR"' ;
              smpDecVirInbqe : SQL:='Select MP_MODEPAIE FROM MODEPAIE WHERE MP_CATEGORIE="TRI"' ;
   END ;
QQ:=OpenSQL(SQL,True) ;
St:='' ;
While Not QQ.EOF do
   BEGIN
   St:=St+'E_MODEPAIE="'+QQ.Fields[0].AsString+'" OR ' ;
   QQ.Next ;
   END ;
Ferme(QQ) ;
if St<>'' then Delete(St,Length(St)-2,3) ;
XX_WHERE1.Text:=St ;
END ;
}
Procedure TFCFONBMP.InitTitres ;
BEGIN
Case smp Of
   smpEnctraEnc : Caption := HM.Mess[14] ;
   smpEncTraEsc : Caption := HM.Mess[13] ;
   smpEncPreBqe : Caption := HM.Mess[15] ;
   smpDecVirbqe : Caption := HM.Mess[16] ;
 smpDecVirInbqe : Caption := HM.Mess[18] ;
   smpDecBorDec : Caption := HM.Mess[17] ;
   END ;
END ;


procedure TFCFONBMP.FormShow(Sender: TObject);
begin
E_DEVISE.Value:=V_PGI.DevisePivot ;
if (smp in [smpDecVirInBqe]) then begin
  E_DEVISE.Enabled := True;
  bRib.Hint := 'Modifier l''IBAN';
  cRIB.Caption := '&IBAN renseigné';
end;
if VH^.CPExoRef.Code<>'' then
   BEGIN
   E_EXERCICE.Value:=VH^.CPExoRef.Code ; E_EXERCICEChange(Nil) ;
   E_DATECOMPTABLE.Text:=DateToStr(VH^.CPExoRef.Deb) ; E_DATECOMPTABLE_.Text:=DateToStr(VH^.CPExoRef.Fin) ;
   END else
   BEGIN
   E_EXERCICE.Value:=VH^.Entree.Code ; E_EXERCICEChange(Nil) ;
   E_DATECOMPTABLE.Text:=DateToStr(V_PGI.DateEntree) ; E_DATECOMPTABLE_.Text:=DateToStr(V_PGI.DateEntree) ;
   END ;
E_DATEECHEANCE.Text:=StDate1900 ; E_DATEECHEANCE_.Text:=StDate2099 ;
inherited;
InitTablesLibresTiers(PLibres) ;
XX_WHERE1.Text:='' ; BanqueGene:='' ;
if isCli then
   BEGIN
   Caption:=HM.Mess[5] ; XX_WHEREDC.Text:='' ; HelpContext:=7586000 ;
   Auxiliaire1.ZoomTable:=tztToutDebit  ; Auxiliaire2.ZoomTable:=tztToutDebit ;
   XX_WHERENATTIERS.Text:='T_NATUREAUXI="CLI" OR T_NATUREAUXI="AUD"' ;
   END else
   BEGIN
   Caption:=HM.Mess[6] ; XX_WHEREDC.Text:='' ; HelpContext:=7595000 ;
   Auxiliaire1.ZoomTable:=tztToutCredit  ; Auxiliaire2.ZoomTable:=tztToutCredit ;
   XX_WHERENATTIERS.Text:='T_NATUREAUXI="FOU" OR T_NATUREAUXI="AUC"' ;
   END ;
InitTitres ;
UpdateCaption(Self) ;
end;

procedure TFCFONBMP.Auxiliaire1Change(Sender: TObject);
begin
E_AUXILIAIRE.Text := Auxiliaire1.Text ;
end;

procedure TFCFONBMP.Auxiliaire2Change(Sender: TObject);
begin
E_AUXILIAIRE_.Text := Auxiliaire2.Text ;
end;

procedure TFCFONBMP.E_EXERCICEChange(Sender: TObject);
begin ExoToDates(E_EXERCICE.Value,E_DATECOMPTABLE,E_DATECOMPTABLE_) ; end;

procedure TFCFONBMP.E_DATECOMPTABLEKeyPress(Sender: TObject; var Key: Char);
begin ParamDate(Self,Sender,Key) ; end;

Function TFCFONBMP.GetLeOBM : TOBM ;
Var Q1 : TQuery ;
    O  : TOBM ;
BEGIN
{Result:=Nil ;} O:=Nil ;
Q1:=OpenSQL('Select * from Ecriture where E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'"'
          +' AND E_EXERCICE="'+QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime))+'"'
          +' AND E_DATECOMPTABLE="'+USDATETIME(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'"'
          +' AND E_NUMEROPIECE='+Q.FindField('E_NUMEROPIECE').AsString
          +' AND E_NUMLIGNE='+Q.FindField('E_NUMLIGNE').AsString
          +' AND E_QUALIFPIECE="N"'
          +' AND E_NUMECHE='+Q.FindField('E_NUMECHE').AsString,True) ;
if Not Q1.EOF then BEGIN O:=TOBM.Create(EcrGen,'',False) ; O.ChargeMvt(Q1) ; END ;
Ferme(Q1) ;
Result:=O ;
END ;

procedure TFCFONBMP.bRibClick(Sender: TObject);
var O  : TOBM ;
    IsAux : Boolean ;
begin
if Q.EOF then Exit ;
O:=GetLeOBM ;
IsAux:=O.GetMvt('E_AUXILIAIRE')<>'' ;
if (smp in [smpDecVirInBqe]) then begin
  if O<>Nil then if ModifRibOBM(O,True,FALSE,'',IsAux,True) then bChercheClick(nil);
  end
else begin
  if O<>Nil then if ModifRibOBM(O,True,FALSE,'',IsAux) then bChercheClick(nil);
end;
O.Free ;
end;

procedure TFCFONBMP.SQDataChange(Sender: TObject; Field: TField);
begin
if not Q.Active then exit ;
bRib.Enabled := not (Q.EOF) ;
end;

procedure TFCFONBMP.cRIBClick(Sender: TObject);
begin
case cRIB.State of
   cbGrayed    : XX_WHERERIB.Text := '' ;
   cbChecked   : XX_WHERERIB.Text := 'E_RIB<>""' ;
   cbUnchecked : XX_WHERERIB.Text := 'E_RIB="" or E_RIB="////"' ;
   end ;
end;

{procedure TFCFONBMP.CompleteTL ( TL : Tlist ) ;
Var i : integer ;
    O,O2 : TOBM ;
    RR   : RMVT ;
    QL : TQuery ;
//    T  : TStrings ;
    CodeL,SQL,St : String ;
BEGIN
for i:=0 to TL.Count-1 do
    BEGIN
    O:=TOBM(TL[i]) ; if O=Nil then Break ;
    CodeL:=O.GetMvt('E_LETTRAGE') ; if CodeL='' then Continue ;
    if O.GetMvt('E_ETATLETTRAGE')='RI' then Continue ;
    SQL:='Select * from Ecriture Where E_AUXILIAIRE="'+O.GetMvt('E_AUXILIAIRE')+'"'
        +'AND E_GENERAL="'+O.GetMvt('E_GENERAL')+'" AND E_LETTRAGE="'+CodeL+'"' ;
    QL:=OpenSQL(SQL,True) ;
    While Not QL.EOF do
       BEGIN
       O2:=TOBM.Create(EcrGen,'',False) ; O2.ChargeMvt(QL) ;
       // Ne pas se reprendre elle-même
       if ((O2.GetMvt('E_NUMEROPIECE')<>O.GetMvt('E_NUMEROPIECE')) or
           (O2.GetMvt('E_NUMLIGNE')<>O.GetMvt('E_NUMLIGNE')) or
           (O2.GetMvt('E_NUMECHE')<>O.GetMvt('E_NUMECHE')) or
           (O2.GetMvt('E_JOURNAL')<>O.GetMvt('E_JOURNAL')) or
           (O2.GetMvt('E_DATECOMPTABLE')<>O.GetMvt('E_DATECOMPTABLE'))) then
           BEGIN
           RR:=OBMToIdent(O2,True) ; St:=EncodeLC(RR) ;
           O2.Free ;
           O.LC.Add(St) ;
           {
           T:=TStringList.Create ; T.Add(St) ;
           O.LC.Assign(T) ; T.Free ;
           }
{           END else
           BEGIN
           O2.Free ;
           END ;
       QL.Next ;
       END ;
    Ferme(QL) ;
    END ;
END ;}

procedure TFCFONBMP.cExportClick(Sender: TObject);
begin
case cExport.State of
   cbGrayed    : XX_WHEREEXPORT.Text := '' ;
   cbChecked   : XX_WHEREEXPORT.Text := 'E_CFONBOK="X"' ;
   cbUnchecked : XX_WHEREEXPORT.Text := 'E_CFONBOK<>"X"' ;
   end ;
end;

{Procedure TFCFONBMP.FlagExportLettre ( TL : Tlist ) ;
Var i : integer ;
    O : TOBM ;
    MM : RMVT ;
BEGIN
for i:=0 to TL.Count-1 do
    BEGIN
    O:=TOBM(TL[i]) ; if O=Nil then Exit ;
    MM:=OBMToIdent(O,True) ;
    ExecuteSQL('UPDATE ECRITURE SET E_CFONBOK="X" Where '+WhereEcriture(tsGene,MM,True)) ;
    END ;
END ;}

Function TFCFONBMP.CoherBanque ( TL : TList ) : boolean ;
Var i : integer ;
    O : TOBM ;
    Jal,OldJal : String3 ;
    Okok : boolean ;
    QQ   : TQuery ;
BEGIN
if isCli then BEGIN Result:=True ; Exit ; END ;
Okok:=True ; OldJal:='' ;
for i:=0 to TL.Count-1 do
    BEGIN
    O:=TOBM(TL[i]) ; if O=Nil then Break ;
    Jal:=O.GetMvt('E_JOURNAL') ;
    if ((OldJal<>'') and (Jal<>OldJal)) then BEGIN Okok:=False ; Break ; END ;
    OldJal:=Jal ;
    END ;
if Not Okok then
   BEGIN
   Okok:=(HM.Execute(12,caption,'')=mrYes) ;
   END else if OldJal<>'' then
   BEGIN
   QQ:=OpenSQL('Select J_CONTREPARTIE from JOURNAL Where J_JOURNAL="'+OldJal+'"',True) ;
   if Not QQ.EOF then BanqueGene:=QQ.Fields[0].AsString ;
   Ferme(QQ) ;
   END ;
Result:=Okok ;
END ;

Function TFCFONBMP.ExporteSelection : boolean ;
Var i,NbLig,ii : integer ;
    TL : TList ;
    O       : TOBM ;
//    SWhere,SQL : String ;
    Inutile : TMSEncaDeca ;
BEGIN
{ii:=0 ;} Result:=False ;
Fillchar(Inutile,SizeOf(Inutile),#0) ;
if Not FListe.AllSelected then
   BEGIN
   NbLig:=Fliste.NbSelected ;
   if NbLig<=0 then BEGIN HM.Execute(0,caption,'') ; Exit ; END ;
   if HM.Execute(1,caption,'')<>mrYes then Exit ;
   TL:=Tlist.Create ;
   for i:=0 to NbLig-1 do
       BEGIN
       Fliste.GotoLeBookMark(i) ;
       O:=GetLeOBM ; if O<>Nil then TL.Add(O) ;
       END ;
   END else
   BEGIN
   if HM.Execute(2,caption,'')<>mrYes then Exit ;
   TL:=Tlist.Create ;
   Q.First ; ii:=0 ;
   While ((Not Q.EOF) and (ii<1000)) do
      BEGIN
      O:=GetLeOBM ; if O<>Nil then TL.Add(O) ;
      Q.Next ; Inc(ii) ;
      END ;
   END ;
if CoherBanque(TL) then ExportCFONB(isCli,BanqueGene,'','DEM',TL,smp) ;
VideListe(TL) ; TL.Free ;
Result:=True ; BChercheClick(Nil) ;
END ;

procedure TFCFONBMP.BOuvrirClick(Sender: TObject);
begin
//  inherited;
if Not ExporteSelection then Exit ;
if Not FListe.AllSelected then Fliste.ClearSelected else FListe.AllSelected:=False ;
end;

procedure TFCFONBMP.FListeDblClick(Sender: TObject);
begin
  inherited;
if Q.EOF then Exit ;
TrouveEtLanceSaisie(Q,taConsult,E_QUALIFPIECE.Text) ;
end;

procedure TFCFONBMP.CTIDTICClick(Sender: TObject);
begin
  inherited;
Auxiliaire1.Visible:=Not CTIDTIC.Checked ; Auxiliaire2.Visible:=Not CTIDTIC.Checked ;
HAuxiliaire1.Visible:=Not CTIDTIC.Checked ; HAuxiliaire2.Visible:=Not CTIDTIC.Checked ;
If CTIDTIC.Checked Then
  BEGIN
  Auxiliaire1.Text:='' ; Auxiliaire2.Text:='' ;
  if isCli Then E_GENERAL.ZoomTable:=tzGTID else E_GENERAL.ZoomTable:=tzGTIC ; E_GENERAL.Text:='' ;
  XX_WHEREAUX.Text:='E_AUXILIAIRE="" AND E_NUMECHE>0 AND E_ETATLETTRAGE<>"RI" ' ;
  END Else
  BEGIN
  if isCli then E_GENERAL.ZoomTable:=tzGCollClient else E_GENERAL.ZoomTable:=tzGCollFourn ;
  XX_WHEREAUX.Text:='E_AUXILIAIRE<>""' ;
  END ;
end;

{ BPY 29/01/2003 Correction de la fiche 10622 }
procedure TFCFONBMP.BCtrlRibClick(Sender: TObject);
var
    StWRib : String ;
    i : Integer;
begin
    inherited;
    StWRib := RecupWhereCritere(Pages) ;
    If (StWRib = '') Then Exit;
    // Si on n'est pas en Sélection inversée
    if ((Not FListe.AllSelected) and (FListe.NbSelected>0) and (FListe.NbSelected<100)) then
    begin  // Si on n'a pas tous sélectionné ET qu'il y a au moins 1 et 100 au plus lignes sélectionnées
        // Rajoute une clause au WHERE
        StWRib := StWRib+' AND (';
        for i:=0 to FListe.NbSelected-1 do
        begin
            FListe.GotoLeBookmark(i) ;
            StWRib := StWRib +' (E_NUMEROPIECE='+ Q.FindField('E_NUMEROPIECE').AsString +' AND E_NUMLIGNE='+ Q.FindField('E_NUMLIGNE').AsString +' AND E_JOURNAL="'+ Q.FindField('E_JOURNAL').AsString +'") OR';
        end;
        // Efface le dernier OR et rajoute ')'
        delete(StWRib,length(StWRib)-2,3);
        StWRib := StWRib +')';
    end;
    If StWRib<>'' Then CPLanceFiche_VerifRib('WHERE='+StWRib);
end;
{ Fin BPY }

end.
