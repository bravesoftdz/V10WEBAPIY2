{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 14/04/2003
Modifié le ... :   /  /
Description .. : Remplacé en eAGL par CPCFONB_TOF.PAS
Mots clefs ... :
*****************************************************************}
unit CFONBAT;

interface

uses
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  Mul,
  HSysMenu,
  Menus,
  DB,
  DBTables,
  Hqry,
  StdCtrls,
  Grids,
  DBGrids,
  HDB,
  ComCtrls,
  HRichEdt,
  Hctrls,
  ExtCtrls,
  Buttons,
  Mask,
  Hcompte,
  Ent1,
  HEnt1,
  ParamDat,
  SaisUtil,
  SaisComm,
  hmsgbox,
  CFONB,
  Saisie,
  LettUtil,
  HTB97,
  ed_tools,
  ColMemo,
  HPanel,
  UiUtil,
  {$IFDEF V530}
  EdtDoc,
  EdtEtat,
  {$ELSE}
  EdtRDoc,
  EdtREtat,
  {$ENDIF}
  DocRegl,
  HRichOLE,
  UTOB,
  uRecupSQLModele
  ;

procedure ExportCFONBBatch(ONB, ENC: boolean; SorteLettre: TSorteLettre; smp: tSuiviMP = smpAucun);

type
  TFMulCFONB = class(TFMul)
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
    H_Document: THLabel;
    Document: THValComboBox;
    Aperc: TCheckBox;
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
    procedure BChercheClick(Sender: TObject); override;
    procedure CTIDTICClick(Sender: TObject);
  private
    BanqueGene: string;
    FormshowEnCours: Boolean;
    smp: TsuiviMp;
    function GetLeOBM: TOBM;
    function ExporteSelection: boolean;
    procedure ChargeWhereTraite;
    procedure CompleteTL(TL: Tlist);
    procedure FlagExportLettre(TL: Tlist);
    function CoherBanque(TL: TList): boolean;
    function GetLeTL(NbLig: integer): TList;
  public
    ONB, ENC: Boolean;
    SorteLettre: TSorteLettre;
  end;


implementation

{$R *.DFM}

uses UtilPgi;

procedure ExportCFONBBatch(ONB, ENC: boolean; SorteLettre: TSorteLettre; smp: tSuiviMP = smpAucun);
var
  X: TFMulCFONB;
  PP: THPanel;
begin
  if _Blocage(['nrCloture', 'nrBatch', 'nrLettrage'], True, 'nrAucun') then Exit;
  PP := FindInsidePanel;
  X := TFMulCFONB.Create(Application);
  X.XX_WHERE1.Text := 'E_JOURNAL="zzz"';
  X.Q.Liste := 'MULCFONBE';
  X.ONB := ONB;
  X.ENC := ENC;
  X.SorteLettre := SorteLettre;
  X.Smp := smp;
  case SorteLettre of
    tslBOR: X.FNomFiltre := 'MULBOR';
    tslTraite: X.FNomFiltre := 'MULTRAITE';
  else X.FNomFiltre := 'MULCFONB';
  end;
  if PP = nil then
  begin
    try
      X.ShowModal;
    finally
      X.Free;
    end;
    Screen.Cursor := SyncrDefault;
  end else
  begin
    InitInside(X, PP);
    X.Show;
  end;
end;

procedure TFMulCFONB.ChargeWhereTraite;
var
  QQ: TQuery;
  St: string;
begin
  QQ := OpenSQL('Select MP_MODEPAIE FROM MODEPAIE WHERE MP_LETTRETRAITE="X"', True);
  St := '';
  while not QQ.EOF do
  begin
    St := St + 'E_MODEPAIE="' + QQ.Fields[0].AsString + '" OR ';
    QQ.Next;
  end;
  Ferme(QQ);
  if St <> '' then Delete(St, Length(St) - 2, 3);
  XX_WHERE1.Text := St;
end;

procedure TFMulCFONB.FormShow(Sender: TObject);
begin
  FormShowEnCours := TRUE;
  E_DEVISE.Value := V_PGI.DevisePivot;
  //if ONB then E_DEVISE.Enabled:=False ; // 14741
  if VH^.CPExoRef.Code <> '' then
  begin
    E_EXERCICE.Value := VH^.CPExoRef.Code;
    E_EXERCICEChange(nil);
    E_DATECOMPTABLE.Text := DateToStr(VH^.CPExoRef.Deb);
    E_DATECOMPTABLE_.Text := DateToStr(VH^.CPExoRef.Fin);
  end else
  begin
    E_EXERCICE.Value := VH^.Entree.Code;
    E_EXERCICEChange(nil);
    E_DATECOMPTABLE.Text := DateToStr(V_PGI.DateEntree);
    E_DATECOMPTABLE_.Text := DateToStr(V_PGI.DateEntree);
  end;
  E_DATEECHEANCE.Text := StDate1900;
  E_DATEECHEANCE_.Text := StDate2099;
  inherited;
  InitTablesLibresTiers(PLibres);
  XX_WHERE1.Text := '';
  BanqueGene := '';
  if SorteLettre <> tslAucun then
  begin
    H_Document.Visible := True;
    Document.Visible := True;
    Aperc.Visible := True;
    Document.DataType := 'TTMODELELETTRETRA';
    ChargeComboEtat(Document,'L','LTR','') ;  //XVI FQ 15425 (non corrigé encore)
    H_Document.Caption := HM.Mess[9];
    E_MODEPAIE.DATATYPE := 'TTMODEPAIETRAITE';
    ChargeWhereTraite;
    Aperc.Visible := False;
    CExport.Checked := True;
    XX_WHERE1.Text := 'E_ETATLETTRAGE<>"RI" AND E_LETTRAGE<>""';
    if SorteLettre = tslBOR then
    begin
      XX_WHERENAT.Text := 'E_NATUREPIECE="OD" OR E_NATUREPIECE="RF" OR E_NATUREPIECE="AF" OR E_NATUREPIECE="OF"';
      XX_WHEREDC.Text := 'E_DEBIT>0 OR E_CREDIT<0';
      Caption := HM.Mess[7];
    end else
    begin
      XX_WHERENAT.Text := 'E_NATUREPIECE="OD" OR E_NATUREPIECE="RC" OR E_NATUREPIECE="AC" OR E_NATUREPIECE="OC"';
      XX_WHEREDC.Text := 'E_DEBIT<0 OR E_CREDIT>0';
      Caption := HM.Mess[11];
    end;
  end else if not ONB then
  begin
    H_Document.Visible := True;
    ChargeComboEtat(Document,'E','BOR','') ; //XVI FQ 15425 (non corrigé encore)
    Document.Visible := True;
    Aperc.Visible := True;
    E_JOURNAL.DataType := '';
    E_JOURNAL.Vide := False;
    E_JOURNAL.DataType := 'TTJALSANSECART';
    if E_JOURNAL.Items.Count > 0 then E_JOURNAL.ItemIndex := 0;
    Caption := HM.Mess[4];
  end else
  begin
    if ENC then
    begin
      Caption := HM.Mess[5];
      XX_WHEREDC.Text := '';
    end
    else
    begin
      Caption := HM.Mess[6];
      XX_WHEREDC.Text := '';
    end;
  end;
  if ONB and ENC then HelpContext := 7586000 else
    if (not ONB) and ENC then HelpContext := 7598210 else // 14371 En attente n° d'aide 999999824 else
    if (ONB) and (not ENC) then HelpContext := 7595000;
  if SorteLettre = tslTraite then HelpContext := 7595200 else
    if SorteLettre = tslBOR then HelpContext := 7595400;
  //If ONB And (SorteLettre=tslAucun) And (Not Document.Visible) Then CTIDTIC.Visible:=TRUE ;
  if (SorteLettre = tslAucun) then CTIDTIC.Visible := TRUE;
  {$IFDEF CCMP}
  if (SorteLettre = tslAucun) and (not ONB) then CTIDTICClick(nil);
  if (SorteLettre = tslAucun) then
  begin
    if ONB and Enc then HelpContext := 7586000;
    if ONB and (not Enc) then HelpContext := 7595000;
    if (not ONB) then HelpContext := 7598210;
  end;
  {$ENDIF}
  UpdateCaption(Self);
  FormShowEnCours := FALSE;
end;

procedure TFMulCFONB.Auxiliaire1Change(Sender: TObject);
begin
  E_AUXILIAIRE.Text := Auxiliaire1.Text;
end;

procedure TFMulCFONB.Auxiliaire2Change(Sender: TObject);
begin
  E_AUXILIAIRE_.Text := Auxiliaire2.Text;
end;

procedure TFMulCFONB.E_EXERCICEChange(Sender: TObject);
begin
  ExoToDates(E_EXERCICE.Value, E_DATECOMPTABLE, E_DATECOMPTABLE_);
  if E_EXERCICE.Value = '' then              //Modif SG6 02/11/2044
  begin
    E_DATECOMPTABLE.Text := StDate1900;
    E_DATECOMPTABLE_.Text := StDate2099;
  end;
end;

procedure TFMulCFONB.E_DATECOMPTABLEKeyPress(Sender: TObject; var Key: Char);
begin
  ParamDate(Self, Sender, Key);
end;

function TFMulCFONB.GetLeOBM: TOBM;
var
  Q1: TQuery;
  O: TOBM;
begin
  {Result:=Nil ;}O := nil;
  Q1 := OpenSQL('Select * from Ecriture where E_JOURNAL="' + Q.FindField('E_JOURNAL').AsString + '"'
    + ' AND E_EXERCICE="' + QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime)) + '"'
    + ' AND E_DATECOMPTABLE="' + USDATETIME(Q.FindField('E_DATECOMPTABLE').AsDateTime) + '"'
    + ' AND E_NUMEROPIECE=' + Q.FindField('E_NUMEROPIECE').AsString
    + ' AND E_NUMLIGNE=' + Q.FindField('E_NUMLIGNE').AsString
    + ' AND E_QUALIFPIECE="N"'
    + ' AND E_NUMECHE=' + Q.FindField('E_NUMECHE').AsString, True);
  if not Q1.EOF then
  begin
    O := TOBM.Create(EcrGen, '', False);
    O.ChargeMvt(Q1);
  end;
  Ferme(Q1);
  Result := O;
end;

procedure TFMulCFONB.bRibClick(Sender: TObject);
var
  O: TOBM;
  IsAux: Boolean;
begin
  if Q.EOF then Exit;
  O := GetLeOBM;
  IsAux := O.GetMvt('E_AUXILIAIRE') <> '';
  if O <> nil then if ModifRibOBM(O, True, FALSE, '', IsAux) then bChercheClick(nil);
  O.Free;
end;

procedure TFMulCFONB.SQDataChange(Sender: TObject; Field: TField);
begin
  if not Q.Active then exit;
  bRib.Enabled := not (Q.EOF);
end;

procedure TFMulCFONB.cRIBClick(Sender: TObject);
begin
  case cRIB.State of
    cbGrayed: XX_WHERERIB.Text := '';
    cbChecked: XX_WHERERIB.Text := 'E_RIB<>""';
    cbUnchecked: XX_WHERERIB.Text := 'E_RIB="" or E_RIB="////"';
  end;
end;

procedure TFMulCFONB.CompleteTL(TL: Tlist);
var
  i: integer;
  O, O2: TOBM;
  RR: RMVT;
  QL: TQuery;
  //    T  : TStrings ;
  CodeL, SQL, St: string;
begin
  for i := 0 to TL.Count - 1 do
  begin
    O := TOBM(TL[i]);
    if O = nil then Break;
    CodeL := O.GetMvt('E_LETTRAGE');
    if CodeL = '' then Continue;
    if O.GetMvt('E_ETATLETTRAGE') = 'RI' then Continue;
    SQL := 'Select * from Ecriture Where E_AUXILIAIRE="' + O.GetMvt('E_AUXILIAIRE') + '"'
      + 'AND E_GENERAL="' + O.GetMvt('E_GENERAL') + '" AND E_LETTRAGE="' + CodeL + '"';
    QL := OpenSQL(SQL, True);
    while not QL.EOF do
    begin
      O2 := TOBM.Create(EcrGen, '', False);
      O2.ChargeMvt(QL);
      // Ne pas se reprendre elle-même
      if ((O2.GetMvt('E_NUMEROPIECE') <> O.GetMvt('E_NUMEROPIECE')) or
        (O2.GetMvt('E_NUMLIGNE') <> O.GetMvt('E_NUMLIGNE')) or
        (O2.GetMvt('E_NUMECHE') <> O.GetMvt('E_NUMECHE')) or
        (O2.GetMvt('E_JOURNAL') <> O.GetMvt('E_JOURNAL')) or
        (O2.GetMvt('E_DATECOMPTABLE') <> O.GetMvt('E_DATECOMPTABLE'))) then
      begin
        RR := OBMToIdent(O2, True);
        St := EncodeLC(RR);
        O2.Free;
        O.LC.Add(St);
        {
        T:=TStringList.Create ; T.Add(St) ;
        O.LC.Assign(T) ; T.Free ;
        }
      end else
      begin
        O2.Free;
      end;
      QL.Next;
    end;
    Ferme(QL);
  end;
end;

procedure TFMulCFONB.cExportClick(Sender: TObject);
begin
  case cExport.State of
    cbGrayed: XX_WHEREEXPORT.Text := '';
    cbChecked: XX_WHEREEXPORT.Text := 'E_CFONBOK="X"';
    cbUnchecked: XX_WHEREEXPORT.Text := 'E_CFONBOK<>"X"';
  end;
end;

procedure TFMulCFONB.FlagExportLettre(TL: Tlist);
var
  i: integer;
  O: TOBM;
  MM: RMVT;
begin
  for i := 0 to TL.Count - 1 do
  begin
    O := TOBM(TL[i]);
    if O = nil then Exit;
    MM := OBMToIdent(O, True);
    ExecuteSQL('UPDATE ECRITURE SET E_CFONBOK="X" Where ' + WhereEcriture(tsGene, MM, True));
  end;
end;

function TFMulCFONB.CoherBanque(TL: TList): boolean;
var
  i: integer;
  O: TOBM;
  Jal, OldJal: String3;
  Okok: boolean;
  QQ: TQuery;
begin
  if ENC then
  begin
    Result := True;
    Exit;
  end;
  Okok := True;
  OldJal := '';
  for i := 0 to TL.Count - 1 do
  begin
    O := TOBM(TL[i]);
    if O = nil then Break;
    Jal := O.GetMvt('E_JOURNAL');
    if ((OldJal <> '') and (Jal <> OldJal)) then
    begin
      Okok := False;
      Break;
    end;
    OldJal := Jal;
  end;
  if not Okok then
  begin
    Okok := (HM.Execute(12, caption, '') = mrYes);
  end else if OldJal <> '' then
  begin
    QQ := OpenSQL('Select J_CONTREPARTIE from JOURNAL Where J_JOURNAL="' + OldJal + '"', True);
    if not QQ.EOF then BanqueGene := QQ.Fields[0].AsString;
    Ferme(QQ);
  end;
  Result := Okok;
end;

// BPY le 17/02/2004 => fonction pour obtenir le TL ....
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/02/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TFMulCFONB.GetLeTL(NbLig: integer): TList;
var
  TL: TList;
  i: integer;
  O: TOBM;
begin
  result := nil;

  // recup de chaque ligne !
  if (not FListe.AllSelected) then
  begin
    if (HM.Execute(1, caption, '') <> mrYes) then exit;

    TL := Tlist.Create;
    for i := 0 to NbLig - 1 do
    begin
      Fliste.GotoLeBookMark(i);
      O := GetLeOBM;
      if (O <> nil) then TL.Add(O);
    end;
  end
  else
  begin
    if (HM.Execute(2, caption, '') <> mrYes) then exit;

    TL := Tlist.Create;
    Q.First;
    while (not Q.EOF) do
    begin
      O := GetLeOBM;
      if (O <> nil) then TL.Add(O);
      Q.Next;
    end;
  end;

  result := TL;
end;

// BPY le 08/10/2003
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 08/10/2003
Modifié le ... : 20/02/2004
Description .. : 2004-02-20 : Suppression de la limite a 1000 Lignes ....
Suite ........ :     en passant par une TOB et un LanceEtatTob
Mots clefs ... :
*****************************************************************}
function TFMulCFONB.ExporteSelection: boolean;
var
  i, NbLig: integer;
  TL: TList;
  TE: TOB;
  SQuery, SWhere, AWhere, SGroup, SOrder: string;
  Inutile: TMSEncaDeca;
begin
  result := false;
  Fillchar(Inutile, SizeOf(Inutile), #0);

  // pas de document selectionné .... out !
  if ((not ONB) and (Document.Value = '')) then
  begin
    if (SorteLettre = tslAucun) then HM.Execute(3, Caption, '')
    else HM.Execute(8, Caption, '');
    exit;
  end;

  // Recup du nb de Ligne
  if (FListe.AllSelected) then
    {$IFDEF EAGLCLIENT}
    NbLig := Q.Detail.Count
      {$ELSE}
    NbLig := Q.RecordCount
      {$ENDIF}
  else NbLig := Fliste.NbSelected;

  // si pas de lignes ...
  if (NbLig <= 0) then
  begin
    HM.Execute(0, caption, '');
    exit;
  end;

  // traitement
  if (ONB) then
  begin
    TL := GetLeTL(NbLig);
    //SG6 07.03.05 FQ15462
    if TL = nil then
    begin
      result := False;
      Exit;
    end;
    if (CoherBanque(TL)) then ExportCFONB(ENC, BanqueGene, '', 'DEM', TL);
    VideListe(TL);
    TL.Free;
  end
  else if (SorteLettre <> tslAucun) then
  begin
    TL := GetLeTL(NbLig);
    CompleteTL(TL);
    if (LanceDocRegl(TL, SorteLettre, Document.Value, '', False, Inutile,
                         TRUE , OnLocale , smpAucun , nil, // paramètre par défaut non utilisé
                         True                              // Mode réédition // FQ 14857 SBO 05/11/2004
                     ) > 0) then FlagExportLettre(TL);
    VideListe(TL);
    TL.Free;
  end
  else
  begin
    // BPY le 16/02/2004 => depassement de la limite a 1000 ligne !
//        TL := GetLeTL(NbLig);
//        SWhere := WhereMulti(TL);
//        LanceEtat('E','BOR',Document.Value,Aperc.Checked,FExport.Checked,False,Nil,SWhere,'',False);
//        VideListe(TL);
//        TL.Free;
    TE := TOB.Create('', nil, -1);
    SQuery := RecupSQLComplet('E', 'BOR', Document.Value, SWhere, SGroup, SOrder);

    // recuperation d'une tob pour chaque line !
    // !!! ATTENTION !!! => La clause group by ne sera pas prise en compte !
    if (not FListe.AllSelected) then
    begin
      if (HM.Execute(1, caption, '') <> mrYes) then exit;
      for i := 0 to NbLig - 1 do
      begin
        Fliste.GotoLeBookMark(i);
        AWhere := ' AND E_JOURNAL="' + Q.FindField('E_JOURNAL').AsString + '" AND E_EXERCICE="' + QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime)) +
          '" AND E_DATECOMPTABLE="' + USDATETIME(Q.FindField('E_DATECOMPTABLE').AsDateTime) + '" AND E_NUMEROPIECE=' + Q.FindField('E_NUMEROPIECE').AsString + ' AND E_NUMLIGNE=' +
          Q.FindField('E_NUMLIGNE').AsString + ' AND E_QUALIFPIECE="N"' + ' AND E_NUMECHE=' + Q.FindField('E_NUMECHE').AsString;
        TE.LoadDetailFromSQL(SQuery + ' ' + SWhere + AWhere, true);
      end;
    end
    else
    begin
      if (HM.Execute(2, caption, '') <> mrYes) then exit;
      Q.First;
      while (not Q.EOF) do
      begin
        AWhere := ' AND E_JOURNAL="' + Q.FindField('E_JOURNAL').AsString + '" AND E_EXERCICE="' + QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime)) +
          '" AND E_DATECOMPTABLE="' + USDATETIME(Q.FindField('E_DATECOMPTABLE').AsDateTime) + '" AND E_NUMEROPIECE=' + Q.FindField('E_NUMEROPIECE').AsString + ' AND E_NUMLIGNE=' +
          Q.FindField('E_NUMLIGNE').AsString + ' AND E_QUALIFPIECE="N"' + ' AND E_NUMECHE=' + Q.FindField('E_NUMECHE').AsString;
        TE.LoadDetailFromSQL(SQuery + ' ' + SWhere + AWhere, true);
        Q.Next;
      end;
    end;

    // trie de la tob !
    if (not (SOrder = '')) then
    begin
      SOrder := copy(SOrder, 10, length(SOrder) - 9); // Supprime le "order by "
      while (not (pos(',', Sorder) = 0)) do Sorder[pos(',', Sorder)] := ';';
      TE.Detail.Sort(Sorder);
    end;

    // lancement de l'etat
    // !!! ATTENTION !!! => ne marche pas s'il y a des bande report dans l'etat !
    LanceEtatTob('E', 'BOR', Document.Value, TE, Aperc.Checked, FExport.Checked, false, nil, '', '', false);
    FreeAndNil(TE);
    // fin BPY
  end;

  result := true;
  BChercheClick(nil);
end;
// Fin BPY

procedure TFMulCFONB.BOuvrirClick(Sender: TObject);
begin
  //  inherited;
  if ((SorteLettre <> tslAucun) and (E_GENERAL.ExisteH <= 0)) then
  begin
    HM.Execute(8, Caption, '');
    if E_GENERAL.CanFocus then E_GENERAL.SetFocus;
    Exit;
  end;
  if not ExporteSelection then Exit;
  if not FListe.AllSelected then Fliste.ClearSelected else FListe.AllSelected := False;
end;

procedure TFMulCFONB.FListeDblClick(Sender: TObject);
begin
  inherited;
  if Q.EOF then Exit;
  TrouveEtLanceSaisie(Q, taConsult, E_QUALIFPIECE.Text);
end;


procedure TFMulCFONB.BChercheClick(Sender: TObject);
begin
  if ((SorteLettre <> tslAucun) and (not FormShowEnCours) and (E_GENERAL.ExisteH <= 0)) then
  begin
    HM.Execute(10, Caption, '');
    if E_GENERAL.CanFocus then E_GENERAL.SetFocus;
    Exit;
  end;
  inherited;
end;

procedure TFMulCFONB.CTIDTICClick(Sender: TObject);
begin
  inherited;
  Auxiliaire1.Visible := not CTIDTIC.Checked;
  Auxiliaire2.Visible := not CTIDTIC.Checked;
  HAuxiliaire1.Visible := not CTIDTIC.Checked;
  HAuxiliaire2.Visible := not CTIDTIC.Checked;
  if CTIDTIC.Checked then
  begin
    Auxiliaire1.Text := '';
    Auxiliaire2.Text := '';
    if ENC then E_GENERAL.ZoomTable := tzGTID else E_GENERAL.ZoomTable := tzGTIC;
    E_GENERAL.Text := '';
    XX_WHEREAUX.Text := 'E_AUXILIAIRE="" AND E_NUMECHE>0 AND E_ETATLETTRAGE<>"RI" '
  end else
  begin
    {$IFDEF CCMP}
    if Enc then
    begin
      E_GENERAL.ZoomTable := tzGCollClient;
      Auxiliaire1.ZoomTable := tzTToutDebit;
      Auxiliaire2.ZoomTable := tzTToutDebit;
    end else
    begin
      E_GENERAL.ZoomTable := tzGCollFourn;
      Auxiliaire1.ZoomTable := tzTToutCredit;
      Auxiliaire2.ZoomTable := tzTToutCredit;
    end;
    {$ELSE}
    E_GENERAL.ZoomTable := tzGCollectif;
    {$ENDIF}
    XX_WHEREAUX.Text := 'E_AUXILIAIRE<>""';
  end;
end;

end.

