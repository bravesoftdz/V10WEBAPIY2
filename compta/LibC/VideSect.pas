unit VideSect;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, Hctrls, Mask, hmsgbox, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  DelVisuE, HStatus,
  Ent1, HEnt1, LettUtil, HQry, SaisUtil, ExtCtrls, SaisComm, LetBatch,
  HSysMenu, HCompte, ParamDat, HFLabel, ed_tools, HPanel, UiUtil, HTB97, 
  Menus, ComCtrls
  {$IFDEF MODENT1}
  , CPTypeCons
  {$ENDIF MODENT1}
  {SG6 22.02.05}
  , UTob
  , Ulibanalytique
  , UObjFiltres
  , UtilSais
  ;

type
  P_RAP = class
    Section, General: String17;
    Sousplan: array[1..MaxAxe] of string17;
    Solde {,SoldeEuro}: double;
  end;
type
  TFVideSect = class(TForm)
    HMCle: THMsgBox;
    HMTrad: THSystemMenu;
    FlashVide: TFlashingLabel;
    BGenere: TToolbarButton97;
    BValide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    HPB: TToolWindow97;
    Dock: TDock97;
    HPanel2: THPanel;
    HPanel1: THPanel;
    Pages: TPageControl;
    FFiltres: THValComboBox;
    BFiltre: TToolbarButton97;
    POPF: TPopupMenu;
    TabSheet1: TTabSheet;
    GCRIT: TGroupBox;
    HLabel3: THLabel;
    HLabel6: THLabel;
    TY_AXE: THLabel;
    TS_CLEREPARTITION: THLabel;
    HLabel1: THLabel;
    HLabel2: THLabel;
    TGENEEMET: THLabel;
    TGENERECEPT: THLabel;
    Y_DATECOMPTABLE: THCritMaskEdit;
    Y_DATECOMPTABLE_: THCritMaskEdit;
    Y_AXE: THValComboBox;
    S_CLEREPARTITION: THValComboBox;
    Y_DATECOMPTABLE1: THCritMaskEdit;
    Y_DATECOMPTABLE1_: THCritMaskEdit;
    GENEEMET: TEdit;
    GENERECEPT: TEdit;
    GPARAM: TGroupBox;
    H_JOURNALE: THLabel;
    HLabel9: THLabel;
    HLabel7: THLabel;
    HLabel10: THLabel;
    Y_JOURNAL: THValComboBox;
    DATEGENERATION: THCritMaskEdit;
    Y_REFINTERNE: TEdit;
    Y_LIBELLE: TEdit;
    VERIFQTE: TCheckBox;
    procedure Y_AXEChange(Sender: TObject);
    procedure DATEGENERATIONExit(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BValideClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BAbandonClick(Sender: TObject);
    procedure BGenereClick(Sender: TObject);
    procedure Y_JOURNALChange(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure Y_DATECOMPTABLEKeyPress(Sender: TObject; var Key: Char);
    procedure S_CLEREPARTITIONChange(Sender: TObject);
  private
    //SG6 24.02.05 FQ 15434
    ObjFiltre : TObjFiltre;
    DateGene: TDateTime;
    NumAxe: integer;
    AxeJal: String3;
    WhereEmet, WhereRecept: string;
    TSource, TDest, TPieces: TList;
    //SG6 23.02.05 Gestion ana croisaxe
    vTobAna : TOB;
    function ControleOK: boolean;
    function FabricSQLAnal(ModeR, QQ, Sect: string; Emet: boolean): string;
    procedure RemplirLaListe(Q: TQuery; ModeR: string; Source: boolean);
    function ChargeDest(ModeR, QQ: string): string;
    procedure GenereLesPieces;
    procedure CreerLigneAna(NewNum, NumV: Longint; Section, General: String17; Debit, Credit, DE, CE: Double; SousPlan: array of string17);
    procedure FinFlash;
    function CoherQte(ModeR, QQ: string): boolean;
    procedure MajDernTIS;
    function DejaTIS: boolean;
    function AnalyseOK: boolean;
    function TesterWhere(St: string): boolean;
  public
  end;

procedure VidageInterSections;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPProcGen,
  CPProcMetier,
  {$ENDIF MODENT1}
  UtilPGI;


procedure VidageInterSections;
var
  X: TFVideSect;
  PP: THPanel;
begin
  if PasCreerDate(V_PGI.DateEntree) then Exit;
  if _Blocage(['nrCloture', 'nrBatch', 'nrSaisieModif'], False, 'nrBatch') then Exit;
  X := TFVideSect.Create(Application);
  PP := FindInsidePanel;
  if PP = nil then
  begin
    try
      X.ShowModal;
    finally
      X.Free;
      _Bloqueur('nrBatch', False);
    end;
    Screen.Cursor := SyncrDefault;
  end else
  begin
    InitInside(X, PP);
    X.Show;
  end;
end;

procedure TFVideSect.BAbandonClick(Sender: TObject);
begin
  Close;
end;

procedure TFVideSect.FinFlash;
begin
  FlashVide.Flashing := False;
  FlashVide.Visible := False;
end;

procedure TFVideSect.BGenereClick(Sender: TObject);
begin
  if TPieces.Count > 0 then VisuPiecesGenere(TPieces, EcrAna, 5);
end;

procedure TFVideSect.FormShow(Sender: TObject);
begin
  DateGeneration.Text := DateToStr(V_PGI.DateEntree);
  Y_DATECOMPTABLE.Text := DateToStr(DebutDeMois(V_PGI.DateEntree));
  Y_DATECOMPTABLE_.Text := DateToStr(FinDeMois(V_PGI.DateEntree));
  Y_DATECOMPTABLE1.Text := DateToStr(DebutDeMois(V_PGI.DateEntree));
  Y_DATECOMPTABLE1_.Text := DateToStr(FinDeMois(V_PGI.DateEntree));
  Y_AXE.Value := 'A1';
  DateGene := V_PGI.DateEntree;
  AxeJal := '';
  //SG6 24.02.05 Gestion Filtre V6
  Pages.TabWidth := Pages.Width - 4;
  ObjFiltre.Charger;
end;

procedure TFVideSect.FormCreate(Sender: TObject);
var
  //SG6 24.02.05 FQ 15343 Gestion filtres
  Composant : TControlFiltre ;
begin
  Composant.Filtre := BFiltre;
  Composant.Filtres := FFiltres;
  //Composant.PopupF := POPF;
  Composant.PageCtrl := Pages;
  ObjFiltre := TObjFiltre.Create(Composant, 'VIDESECT');

  PopUpMenu := ADDMenuPop(PopUpMenu, '', '');
  TSource := TList.Create;
  TDest := TList.Create;
  TPieces := TList.Create;
end;

procedure TFVideSect.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  VideListe(TSource);
  TSource.Free;
  VideListe(TDest);
  TDest.Free;
  VideListe(TPieces);
  TPieces.Free;
  //SG6 24.02.05 Gestion Filtres V6 15434
  ObjFiltre.Free;
  {if Parent is THPanel then
     BEGIN
     _Bloqueur('nrBatch',False) ;
     Action:=caFree ;
     END ;}

    //SG6 22.02.05
  if IsInside(Self) then
  begin
    _Bloqueur('nrBatch', False);
    CloseInsidePanel(Self);
  end;
end;

{============================ GESTION ENTETE ET PARAMETRES =================================}
procedure TFVideSect.Y_AXEChange(Sender: TObject);
begin
  if Y_AXE.Value = 'A1' then
  begin
    S_CLEREPARTITION.DataType := 'ttCleRepart1';
    NumAxe := 1
  end else
;    if Y_AXE.Value = 'A2' then
  begin
    S_CLEREPARTITION.DataType := 'ttCleRepart2';
    NumAxe := 2;
  end else
    if Y_AXE.Value = 'A3' then
  begin
    S_CLEREPARTITION.DataType := 'ttCleRepart3';
    NumAxe := 3;
  end else
    if Y_AXE.Value = 'A4' then
  begin
    S_CLEREPARTITION.DataType := 'ttCleRepart4';
    NumAxe := 4;
  end else
    if Y_AXE.Value = 'A5' then
  begin
    S_CLEREPARTITION.DataType := 'ttCleRepart5';
    NumAxe := 5;
  end;
  //SG6 22.02.05 Gestion axe croisé
  if not VH^.AnaCroisaxe then
  begin
    if Y_Axe.Value <> AxeJal then
    begin
      AxeJal := '';
      Y_Journal.Value := '';
    end;
  end;

end;

procedure TFVideSect.DATEGENERATIONExit(Sender: TObject);
var
  DD: TDateTime;
  Err: integer;
begin
  if not IsValidDate(DATEGENERATION.Text) then
  begin
    HMCle.Execute(0, '', '');
    DATEGENERATION.Text := DateToStr(V_PGI.DateEntree);
    DateGene := V_PGI.DateEntree;
  end else
  begin
    DD := StrToDate(DATEGENERATION.Text);
    Err := DateCorrecte(DD);
    if Err > 0 then
    begin
      HMCle.Execute(Err, '', '');
      DATEGENERATION.Text := DateToStr(V_PGI.DateEntree);
      DateGene := V_PGI.DateEntree;
    end else
    begin
      if RevisionActive(DD) then
      begin
        DATEGENERATION.Text := DateToStr(V_PGI.DateEntree);
        DateGene := V_PGI.DateEntree;
      end else
      begin
        DateGene := DD;
      end;
    end;
  end;
end;

function TFVideSect.TesterWhere(St: string): boolean;
var
  Q: TQuery;
  SQL: string;
begin
  Result := True;
  if St = '' then Exit;
  SQL := 'Select * from ANALYTIQ Where ' + St;
  try
    Q := OpenSQL(SQL, True);
  except
    Result := False;
  end;
  Ferme(Q);
end;

function TFVideSect.AnalyseOK: boolean;
var
  St: string;
begin
  Result := False;
  St := GeneEmet.Text;
  if St <> '' then
  begin
    WhereEmet := AnalyseCompte(St, fbGene, False, False);
    WhereEmet := FindEtReplace(WhereEmet, 'G_GENERAL', 'Y_GENERAL', True);
  end;
  St := GeneRecept.Text;
  if St <> '' then
  begin
    WhereRecept := AnalyseCompte(St, fbGene, False, False);
    WhereRecept := FindEtReplace(WhereRecept, 'G_GENERAL', 'Y_GENERAL', True);
  end;
  if not TesterWhere(WhereEmet) then Exit;
  if not TesterWhere(WhereRecept) then Exit;
  Result := True;
end;

function TFVideSect.ControleOK: boolean;
begin
  Result := False;
  if VH^.Cpta[AxeToFb(Y_AXE.Value)].AxGenAttente = '' then
  begin
    HMCle.Execute(18, '', '');
    Exit;
  end;
  if S_CLEREPARTITION.Value = '' then
  begin
    HMCle.Execute(5, '', '');
    Exit;
  end;
  if Y_JOURNAL.Value = '' then
  begin
    HMCle.Execute(6, '', '');
    Exit;
  end;
  if Y_AXE.Value = '' then
  begin
    HMCle.Execute(7, '', '');
    Exit;
  end;
  if not AnalyseOK then
  begin
    HMCle.Execute(20, '', '');
    Exit;
  end;
  Result := True;
end;

{======================================= REQUETE ===========================================}
function TFVideSect.FabricSQLAnal(ModeR, QQ, Sect: string; Emet: boolean): string;
var
  St, StW, StXP, StXN: string;
begin
  if Sect = '' then
  begin
    StW := ' Y_DATECOMPTABLE>="' + USDATE(Y_DATECOMPTABLE) + '"'
      + ' AND Y_DATECOMPTABLE<="' + USDATE(Y_DATECOMPTABLE_) + '"'
      + ' AND S_CLEREPARTITION="' + S_CLEREPARTITION.Value + '"';
  end else
  begin
    StW := ' S_SECTION="' + Sect + '"'
      + ' AND Y_DATECOMPTABLE>="' + USDATE(Y_DATECOMPTABLE1) + '"'
      + ' AND Y_DATECOMPTABLE<="' + USDATE(Y_DATECOMPTABLE1_) + '"';
    if ModeR = 'MC' then ModeR := 'MT'; {pas de subdivision pour les sections réceptrices}
  end;

  StW := StW + ' AND (Y_DEBIT+Y_CREDIT)<>0 AND Y_QUALIFPIECE="N" ';

  if ((Emet) and (WhereEmet <> '')) then StW := StW + ' AND ' + WhereEmet;
  if ((not Emet) and (WhereRecept <> '')) then StW := StW + ' AND ' + WhereRecept;

  StXN := StrFPoint(-9 * Resolution(V_PGI.OkDecV + 1 ));
  StXP := StrFPoint(9 * Resolution(V_PGI.OkDecV + 1 ));

  if ModeR = 'MT' then
  begin

    //SG6 22.02.05 Gestion ana croisaxe
    if VH^.AnaCroisaxe then
      St := 'Select S_SECTION, Sum(Y_DEBIT-Y_CREDIT), Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5'
        + ' From SECTION Left Join ANALYTIQ ON S_SECTION=Y_SOUSPLAN' + Y_AXE.Value[2]
        + ' Where ' + StW + ' Group By S_SECTION, Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5 HAVING Sum(Y_DEBIT-Y_CREDIT) Not Between ' + StXN +
          ' AND ' + StXP
    else
      St := 'Select S_SECTION, Sum(Y_DEBIT-Y_CREDIT)'
        + ' From SECTION Left Join ANALYTIQ ON S_AXE=Y_AXE AND S_SECTION=Y_SECTION'
        + ' Where ' + StW + ' Group By S_SECTION HAVING Sum(Y_DEBIT-Y_CREDIT) Not Between ' + StXN + ' AND ' + StXP;

  end
  else if ModeR = 'MC' then
  begin
    //SG6 22.02.05 Gestion ana croisaxe
    if VH^.AnaCroisaxe then
      St := 'Select S_SECTION, Sum(Y_DEBIT-Y_CREDIT), Y_GENERAL, Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5'
        + ' From SECTION Left Join ANALYTIQ ON S_SECTION=Y_SOUSPLAN' + Y_AXE.Value[2]
        + ' Where ' + StW + ' Group By S_SECTION, Y_GENERAL, Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5 HAVING Sum(Y_DEBIT-Y_CREDIT) Not Between '
          + StXN + ' AND ' + StXP
    else
      St := 'Select S_SECTION, Sum(Y_DEBIT-Y_CREDIT), Y_GENERAL'
        + ' From SECTION Left Join ANALYTIQ ON S_AXE=Y_AXE AND S_SECTION=Y_SECTION'
        + ' Where ' + StW + ' Group By S_SECTION, Y_GENERAL HAVING Sum(Y_DEBIT-Y_CREDIT) Not Between ' + StXN + ' AND ' + StXP;

  end
  else if ModeR = 'QT1' then
  begin

    //SG6 22.02.05 Gestion ana croisaxe
    if VH^.AnaCroisaxe then
      St := 'Select S_SECTION, Sum(Y_DEBIT-Y_CREDIT), Sum(Y_QTE1*(Y_DEBIT-Y_CREDIT)/(Y_DEBIT+Y_CREDIT)), Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5'
        + ' From SECTION Left Join ANALYTIQ S_SECTION=Y_SOUSPLAN' + Y_AXE.Value[2]
        + ' Where ' + StW
    else
      St := 'Select S_SECTION, Sum(Y_DEBIT-Y_CREDIT), Sum(Y_QTE1*(Y_DEBIT-Y_CREDIT)/(Y_DEBIT+Y_CREDIT))'
        + ' From SECTION Left Join ANALYTIQ ON S_AXE=Y_AXE AND S_SECTION=Y_SECTION'
        + ' Where ' + StW;

    if QQ <> '...' then St := St + ' AND Y_QUALIFQTE1="' + QQ + '"';
    if Sect = '' then
    begin
      //SG6 23.02.05 Gestion mode croisaxe
      if VH^.AnaCroisaxe then
        St := St + ' Group By S_SECTION, Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5 HAVING Sum(Y_DEBIT-Y_CREDIT) Not Between ' + StXN + ' AND ' +
          StXP
      else
        St := St + ' Group By S_SECTION HAVING Sum(Y_DEBIT-Y_CREDIT) Not Between ' + StXN + ' AND ' + StXP;

    end else
    begin
      //SG6 23.02.05 Gestion mode croisaxe
      if VH^.AnaCroisaxe then
        St := St +
          ' Group By S_SECTION, Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5 HAVING Sum(Y_QTE1*(Y_DEBIT-Y_CREDIT)/(Y_DEBIT+Y_CREDIT)) Not Between ' + StXN
          + ' AND ' + StXP
      else
        St := St + ' Group By S_SECTION HAVING Sum(Y_QTE1*(Y_DEBIT-Y_CREDIT)/(Y_DEBIT+Y_CREDIT)) Not Between ' + StXN + ' AND ' + StXP;

    end;
  end
  else if ModeR = 'QT2' then
  begin
    //SG6 22.02.05 Gestion croisaxe
    if VH^.AnaCroisaxe then
      St := 'Select S_SECTION, Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5 Sum(Y_DEBIT-Y_CREDIT), Sum(Y_QTE2*(Y_DEBIT-Y_CREDIT)/(Y_DEBIT+Y_CREDIT))'
        + ' From SECTION Left Join ANALYTIQ ON S_SECTION=Y_SOUSPLAN' + Y_AXE.Value[2]
        + ' Where ' + StW
    else
      St := 'Select S_SECTION, Sum(Y_DEBIT-Y_CREDIT), Sum(Y_QTE2*(Y_DEBIT-Y_CREDIT)/(Y_DEBIT+Y_CREDIT))'
        + ' From SECTION Left Join ANALYTIQ ON S_AXE=Y_AXE AND S_SECTION=Y_SECTION'
        + ' Where ' + StW;

    if QQ <> '...' then St := St + ' AND Y_QUALIFQTE2="' + QQ + '"';
    if Sect = '' then
    begin
      //SG6 23.02.05 Gestion ana croisaxe
      if VH^.AnaCroisaxe then
        St := St + ' Group By S_SECTION, Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5 HAVING Sum(Y_DEBIT-Y_CREDIT) Not Between ' + StXN + ' AND ' +
          StXP
      else
        St := St + ' Group By S_SECTION HAVING Sum(Y_DEBIT-Y_CREDIT) Not Between ' + StXN + ' AND ' + StXP;

    end else
    begin
      //SG6 23.02.05 Gestion ana croisaxe
      if VH^.AnaCroisaxe then
        St := St +
          ' Group By S_SECTION, Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5 HAVING Sum(Y_QTE1*(Y_DEBIT-Y_CREDIT)/(Y_DEBIT+Y_CREDIT)) Not Between ' + StXN
          + ' AND ' + StXP
      else
        St := St + ' Group By S_SECTION HAVING Sum(Y_QTE1*(Y_DEBIT-Y_CREDIT)/(Y_DEBIT+Y_CREDIT)) Not Between ' + StXN + ' AND ' + StXP;

    end;
  end;

  Result := St;
end;

{===================================== TRAITEMENT ==========================================}
procedure TFVideSect.RemplirLaListe(Q: TQuery; ModeR: string; Source: boolean);
var
  X: P_RAP;
begin
  X := P_RAP.Create;
  X.Section := Q.FindField('S_SECTION').AsString;

  if Source then
  begin
    //SG6 22.02.05 Gestion ana croisaxe
    if VH^.AnaCroisaxe then
    begin
      X.Sousplan[1] := Q.FindField('Y_SOUSPLAN1').AsString;
      X.Sousplan[2] := Q.FindField('Y_SOUSPLAN2').AsString;
      X.Sousplan[3] := Q.FindField('Y_SOUSPLAN3').AsString;
      X.Sousplan[4] := Q.FindField('Y_SOUSPLAN4').AsString;
      X.Sousplan[5] := Q.FindField('Y_SOUSPLAN5').AsString;
    end;

    X.Solde := Q.Fields[1].AsFloat;
    //X.SoldeEuro:=Q.Fields[2].AsFloat ;
    if ModeR = 'MC' then X.General := Q.Fields[2].AsString;
  end else
  begin
    if ((ModeR = 'QT1') or (ModeR = 'QT2')) then
    begin
      X.Solde := Q.Fields[2].AsFloat;
      //X.SoldeEuro:=Q.Fields[2].AsFloat ;
    end else
    begin
      X.Solde := Q.Fields[1].AsFloat;
      //X.SoldeEuro:=Q.Fields[2].AsFloat ;
    end;
  end;
  if Source then TSource.Add(X) else TDest.Add(X);
end;

function TFVideSect.CoherQte(ModeR, QQ: string): boolean;
var
  St, StW, Section: string;
  Q, QV: TQuery;
  Trouv: boolean;
begin
  Result := True;
  {Sortie triviale}
  if not VerifQTE.Checked then Exit;
  if ((ModeR <> 'QT1') and (ModeR <> 'QT2')) then Exit;
  if QQ = '...' then Exit;
  {Section emettrices}
  StW := ' Y_DATECOMPTABLE>="' + USDATE(Y_DATECOMPTABLE) + '" AND Y_DATECOMPTABLE<="' + USDATE(Y_DATECOMPTABLE_) + '"'
    + ' AND S_CLEREPARTITION="' + S_CLEREPARTITION.Value + '" AND Y_QUALIFPIECE="N"';
  if ModeR = 'QT1' then StW := StW + ' AND Y_QUALIFQTE1<>"' + QQ + '" AND Y_QTE1<>0'
  else StW := StW + ' AND Y_QUALIFQTE2<>"' + QQ + '" AND Y_QTE2<>0';

  //SG6 22.02.5 Gestion Croisaxe
  if VH^.AnaCroisaxe then
    St := 'Select S_SECTION From SECTION Left Join ANALYTIQ ON S_SECTION=Y_SOUSPLAN' + Y_AXE.Value[2] + ' Where ' + StW
  else
    St := 'Select S_SECTION From SECTION Left Join ANALYTIQ ON S_AXE=Y_AXE AND S_SECTION=Y_SECTION Where ' + StW;

  Q := OpenSQL(St, True);
  Trouv := not Q.EOF;
  Ferme(Q);
  if Trouv then
  begin
    Result := False;
    Exit;
  end;

  {Section receptrices}
  StW := ' Y_DATECOMPTABLE>="' + USDATE(Y_DATECOMPTABLE1) + '" AND Y_DATECOMPTABLE<="' + USDATE(Y_DATECOMPTABLE1_) + '" AND Y_QUALIFPIECE="N" ';
  if ModeR = 'QT1' then StW := StW + ' AND Y_QUALIFQTE1<>"' + QQ + '" AND Y_QTE1<>0'
  else StW := StW + ' AND Y_QUALIFQTE2<>"' + QQ + '" AND Y_QTE2<>0';
  QV := OpenSQL('SELECT V_SECTION FROM VENTIL WHERE V_NATURE="CL' + IntToStr(NumAxe) + '"'
    + 'AND V_COMPTE="' + S_CLEREPARTITION.Value + '" ORDER BY V_NATURE, V_COMPTE, V_NUMEROVENTIL', True);
  if QV.EOF then
  begin
    Ferme(QV);
    Exit;
  end;
  while not QV.EOF do
  begin
    Section := QV.FindField('V_SECTION').AsString;
    if Section = '' then Continue;

    //SG6 22.02.5 Gestion Croisaxe
    if VH^.AnaCroisaxe then
      St := 'Select S_SECTION From SECTION Left Join ANALYTIQ ON S_SECTION=Y_SOUSPLAN'
        + Y_AXE.Value[2] + ' Where ' + StW + ' AND S_SECTION="' + Section + '"'
    else
      St := 'Select S_SECTION From SECTION Left Join ANALYTIQ ON S_AXE=Y_AXE AND S_SECTION=Y_SECTION'
        + ' Where ' + StW + ' AND S_SECTION="' + Section + '"';


    Q := OpenSQL(St, True);
    Trouv := not Q.EOF;
    Ferme(Q);
    if Trouv then
    begin
      Result := False;
      Ferme(QV);
      Exit;
    end;
    QV.Next;
  end;
  Ferme(QV);
end;

procedure TFVideSect.MajDernTIS;
var
  Q: TQuery;
begin
  Q := OpenSQL('Select * from EDTLEGAL WHERE ED_OBLIGATOIRE="#"', False);
  Q.Insert;
  InitNew(Q);
  Q.FindField('ED_OBLIGATOIRE').AsString := '-';
  Q.FindField('ED_TYPEEDITION').AsString := 'TIS';
  Q.FindField('ED_EXERCICE').AsString := Y_AXE.Value;
  Q.FindField('ED_PERIODE').AsDateTime := NowH;
  Q.FindField('ED_DATEEDITION').AsDateTime := Date;
  Q.FindField('ED_DATE1').AsDateTime := StrToDate(Y_DATECOMPTABLE.Text);
  Q.FindField('ED_DATE2').AsDateTime := StrToDate(Y_DATECOMPTABLE_.Text);
  Q.FindField('ED_UTILISATEUR').AsString := V_PGI.User;
  Q.FindField('ED_DESTINATION').AsString := S_CLEREPARTITION.Value;
  Q.Post;
  Ferme(Q);
end;

function TFVideSect.DejaTIS: boolean;
var
  Q: TQuery;
  SQL: string;
begin
  Result := False;
  SQL := 'SELECT COUNT(ED_OBLIGATOIRE) from EDTLEGAL Where ED_OBLIGATOIRE="-" AND ED_TYPEEDITION="TIS" AND ED_EXERCICE="' + Y_AXE.Value + '" '
    + 'AND ED_DESTINATION="' + S_CLEREPARTITION.Value + '"';
  SQL := SQL + ' AND ((ED_DATE1<="' + USDATE(Y_DATECOMPTABLE) + '" AND ED_DATE2>="' + USDATE(Y_DATECOMPTABLE) + '")'
    + ' OR (ED_DATE1<="' + USDATE(Y_DATECOMPTABLE_) + '" AND ED_DATE2>="' + USDATE(Y_DATECOMPTABLE_) + '"))';
  Q := OpenSQl(SQL, True);
  if not Q.EOF then if Q.Fields[0].AsInteger >= 1 then Result := True;
  Ferme(Q);
  if Result then
  begin
    if HMCle.Execute(17, '', '') = mrYes then Result := False;
  end;
end;

procedure TFVideSect.BValideClick(Sender: TObject);
var
  St, QQ: string;
  Q, QY: TQuery;
  ModeR: string;
  Tot: double;
  i: integer;
begin
  if not ControleOK then Exit;
  if HMCle.Execute(11, '', '') <> mrYes then Exit;
  if DejaTIS then Exit;
  FlashVide.Visible := True;
  FlashVide.Flashing := True;
  SourisSablier;

  {Lecture des infos de la clef de répartition}
  Q := OpenSQL('Select RE_MODEREPARTITION, RE_QUALIFQTE from CLEREPAR Where '
    + 'RE_AXE="' + Y_AXE.Value + '" AND RE_CLE="' + S_CLEREPARTITION.Value + '"', True);
  if not Q.EOF then
  begin
    {Construire et executer la requête}
    ModeR := Q.Fields[0].AsString;
    QQ := Q.Fields[1].AsString;
    Ferme(Q);
    {Vérification des qtés}
    if not CoherQte(ModeR, QQ) then if HMCle.Execute(13, '', '') <> mrYes then
      begin
        FinFlash;
        Exit;
      end;
    St := FabricSQLAnal(ModeR, QQ, '', True);
  end else
  begin
    Ferme(Q);
    FinFlash;
    Exit;
  end;

  VideListe(TSource);
  VideListe(TDest);

  {Lecture de la requete et chargement des infos emettrices}
  QY := OpenSQL(St, True);
  while not QY.EOF do
  begin
    RemplirLaListe(QY, ModeR, True);
    QY.Next;
  end;
  Ferme(QY);

  if TSource.Count <= 0 then
  begin
    FinFlash;
    HMCle.Execute(8, '', '');
    Exit;
  end;

  {Lecture de la requete et chargement des infos receptrices}
  ChargeDest(ModeR, QQ);
  if TDest.Count <= 0 then
  begin
    FinFlash;
    HMCle.Execute(9, '', '');
    Exit;
  end;
  Tot := 0;
  for i := 0 to TDest.Count - 1 do Tot := Tot + P_RAP(TDest[i]).Solde;
  if Tot = 0 then
  begin
    FinFlash;
    HMCle.Execute(10, '', '');
    Exit;
  end;

  {Génération des pièces}
  if Transactions(GenereLesPieces, 10) <> oeOK then MessageAlerte(HMCle.Mess[14]) else
  begin
    if Transactions(MajDernTIS, 2) <> oeOk then MessageAlerte(HMCle.Mess[16]);
    if TPieces.Count > 0 then
    begin
      if HMCle.Execute(12, '', '') = mrYes then VisuPiecesGenere(TPieces, EcrAna, 5);
    end;
  end;
  SourisNormale;

end;

function TFVideSect.ChargeDest(ModeR, QQ: string): string;
var
  QV, QDest: TQuery;
  St, Section: string;
begin
  QV := OpenSQL('SELECT V_SECTION FROM VENTIL WHERE V_NATURE="CL' + IntToStr(NumAxe) + '"'
    + 'AND V_COMPTE="' + S_CLEREPARTITION.Value + '" ORDER BY V_NATURE, V_COMPTE, V_NUMEROVENTIL', True);
  if QV.EOF then
  begin
    Ferme(QV);
    Exit;
  end;
  while not QV.EOF do
  begin
    Section := QV.FindField('V_SECTION').AsString;
    if Section = '' then Continue;
    St := FabricSQLAnal(ModeR, QQ, Section, False);
    QDest := OpenSQL(St, True);
    while not QDest.EOF do
    begin
      RemplirLaListe(QDest, ModeR, False);
      QDest.Next;
    end;
    Ferme(QDest);
    QV.Next;
  end;
  Ferme(QV);
end;

{================================== FABRICATION PIECES =====================================}
procedure TFVideSect.CreerLigneAna(NewNum, NumV: Longint; Section, General: String17; Debit, Credit, DE, CE: Double; SousPlan: array of string17);
var
  premier_axe : byte;
  i : integer;
  lTOB : TOB; //Lek 271005 FQ16833
begin
  premier_axe := RecherchePremDerAxeVentil.premier_axe;
  if General = '' then General := VH^.Cpta[AxeToFb(Y_AXE.Value)].AxGenAttente;
  vTobAna.PutValue('Y_GENERAL', General);
  if VH^.AnaCroisaxe then
    vTobAna.PutValue('Y_AXE','A' + IntToStr(premier_axe) )
  else
    vTobAna.PutValue('Y_AXE',Y_AXE.Value);

  vTobAna.PutValue('Y_DATECOMPTABLE', DateGene);
  vTobAna.PutValue('Y_PERIODE', GetPeriode(DateGene) );
  vTobAna.PutValue('Y_SEMAINE', NumSemaine(DateGene) );
  vTobAna.PutValue('Y_NUMLIGNE', 0 );
  vTobAna.PutValue('Y_NUMEROPIECE', NewNum );
  vTobAna.PutValue('Y_EXERCICE', QuelExo(DateGeneration.Text) );
  vTobAna.PutValue('Y_REFINTERNE', Copy(Y_REFINTERNE.Text, 1, 35) );
  vTobAna.PutValue('Y_LIBELLE', Copy(Y_LIBELLE.Text, 1, 35) );
  vTobAna.PutValue('Y_QUALIFPIECE', 'N' );
  vTobAna.PutValue('Y_NATUREPIECE', 'OD' );
  vTobAna.PutValue('Y_DEBIT', Debit );
  vTobAna.PutValue('Y_CREDIT', Credit );
  vTobAna.PutValue('Y_DEBITDEV', Debit );
  vTobAna.PutValue('Y_CREDITDEV', Credit );
  vTobAna.PutValue('Y_ETABLISSEMENT', VH^.EtablisDefaut );
  vTobAna.PutValue('Y_DEVISE', V_PGI.DevisePivot );
  vTobAna.PutValue('Y_TAUXDEV', 1 );
  vTobAna.PutValue('Y_TYPEANALYTIQUE', 'X' );
  vTobAna.PutValue('Y_DATEAUXDEV', DateGene );
  vTobAna.PutValue('Y_TOTALECRITURE', 0 );
  vTobAna.PutValue('Y_TOTALDEVISE', 0 );
  vTobAna.PutValue('Y_JOURNAL', Y_JOURNAL.Value );
  vTobAna.PutValue('Y_NUMVENTIL', NumV );
  if NumV = 1 then
    vTobAna.PutValue('Y_TYPEMVT', 'AE' )
  else
    vTobAna.PutValue('Y_TYPEMVT', 'AL' );
  vTobAna.PutValue('Y_CONTROLE', '-' );
  vTobAna.PutValue('Y_ECRANNOUVEAU', 'N' );
  vTobAna.PutValue('Y_CREERPAR', 'GEN' );
  vTobAna.PutValue('Y_EXPORTE', '---' );
  vTobAna.PutValue('Y_QUALIFCRQTE1', '...' );
  vTobAna.PutValue('Y_QUALIFCRQTE2', '...' );

  //Gestion Ana croisaxe
  vTobAna.PutValue('Y_SOUSPLAN' + Y_AXE.Value[2], Section );
  if VH^.AnaCroisaxe then
  begin
    for i := 1 to MaxAxe do
    begin
      if IntToStr(i) = Y_AXE.Value[2] then Continue;
      vTobAna.PutValue('Y_SOUSPLAN' + IntToStr(i), SousPlan[i-1] );
    end;
    vTobAna.PutValue('Y_SECTION', vTobAna.GetString('Y_SOUSPLAN' + IntToStr(premier_axe)) );
  end
  else
  begin
    for i := 1 to MaxAxe do
    begin
      vTobAna.PutValue('Y_SOUSPLAN' + IntToStr(i), '');
    end;
    vTobAna.PutValue('Y_SECTION', Section);
  end;

  vTobAna.InsertDB(nil);

  if ((NumV = 1) and (V_PGI.IoError = oeOk)) then
  begin
    lTOB:= TOB.Create('ANALYTIQ', nil, -1); //Lek 271005 FQ16833
    lTOB.Assign(vTobAna);                   //Lek 271005 FQ16833
    TPieces.Add(lTOB);                      //Lek 271005 FQ16833
  end;
end;

procedure TFVideSect.GenereLesPieces;
var
  isource, idest, NumV: integer;
  Facturier: String3;
  Q: TQuery;
  NewNum: Longint;
  Total, Montant, Debit, Credit, DE, CE, TotDest {,TotalEuro,MontantEuro}: Double;
  Section, SectDest, General, GeneDest: String17;
  X, Z: P_RAP;
  //SGG estion ana croisaxe
begin
  InitMove(TSource.Count, ' ');
  Q := OpenSQL('Select J_COMPTEURNORMAL from JOURNAL Where J_JOURNAL="' + Y_JOURNAL.Value + '"', True);
  Facturier := Q.Fields[0].AsString;
  Ferme(Q);
  vTobAna := TOB.Create('ANALYTIQ', nil, -1);
  TotDest := 0;
  for iDest := 0 to TDest.Count - 1 do
  begin
    Z := P_RAP(TDest[iDest]);
    TotDest := TotDest + Z.Solde;
  end;
  for isource := 0 to TSource.Count - 1 do
  begin
    if MoveCur(False) then ;
    SetIncNum(EcrAna, Facturier, NewNum, DateGene);
    X := P_RAP(TSource[isource]);
    Section := X.Section;
    General := X.General;
    Debit := 0;
    Credit := 0;
    DE := 0;
    CE := 0;
    if VH^.MontantNegatif then
    begin
      if X.Solde < 0 then
      begin
        Credit := X.Solde;
        CE := 0 {X.SoldeEuro};
      end
      else
      begin
        Debit := -X.Solde;
        DE := 0 {-X.SoldeEuro};
      end;
    end else
    begin
      if X.Solde > 0 then
      begin
        Credit := X.Solde;
        CE := 0 {X.SoldeEuro};
      end
      else
      begin
        Debit := -X.Solde;
        DE := 0 {-X.SoldeEuro};
      end;
    end;
    NumV := 1;
    //SG6 22.02.05 Gestion ana croisaxe (rajout param Sousplan)
    CreerLigneAna(NewNum, NumV, Section, General, Debit, Credit, DE, CE, X.Sousplan);
    MajSoldeSectionTOB(vTobAna, True);
    Total := 0; //TotalEuro:=0 ;
    for iDest := 0 to TDest.Count - 1 do
    begin
      Z := P_RAP(TDest[iDest]);
      SectDest := Z.Section;
      GeneDest := Z.General;
      Debit := 0;
      Credit := 0;
      DE := 0;
      CE := 0;
      if ((iDest < TDest.Count - 1) and (TotDest <> 0)) then
      begin
        Montant := Arrondi(X.Solde * Z.Solde / TotDest, V_PGI.OkDecV);
        Total := Total + Montant;
        //MontantEuro:=Arrondi(X.SoldeEuro*Z.Solde/TotDest,V_PGI.OkDecE) ; TotalEuro:=TotalEuro+MontantEuro ;
      end else
      begin
        Montant := Arrondi(X.Solde - Total, V_PGI.OkDecV);
        //MontantEuro:=Arrondi(X.SoldeEuro-TotalEuro,V_PGI.OkDecE) ;
      end;
      if Montant = 0 then Continue;
      if X.Solde > 0 then
      begin
        if ((Montant > 0) or (VH^.MontantNegatif)) then
        begin
          Debit := Montant;
          DE := 0 {MontantEuro};
        end
        else
        begin
          Credit := Abs(Montant);
          CE := 0 {Abs(MontantEuro)};
        end;
      end else
      begin
        if ((Montant < 0) or (VH^.MontantNegatif)) then
        begin
          Credit := -Montant;
          CE := 0 {-MontantEuro};
        end
        else
        begin
          Debit := Montant;
          DE := 0 {MontantEuro};
        end;
      end;
      Inc(NumV);
      CreerLigneAna(NewNum, NumV, SectDest, General, Debit, Credit, DE, CE, X.Sousplan); {gene = celui de origine}
      MajSoldeSectionTOB(vTobAna, True);
    end;
  end;
  FreeAndNil(vTobAna); //Lek 271005 je remets mais je le crée pour TPices FQ16833
  FinFlash;
  FiniMove;
end;

procedure TFVideSect.Y_JOURNALChange(Sender: TObject);
var
  Q: TQuery;
begin
  AxeJal := '';
  if Y_Journal.Value = '' then Exit;
  Q := OpenSQL('Select J_AXE, J_COMPTEURNORMAL from JOURNAL Where J_JOURNAL="' + Y_Journal.Value + '"', True);
  if not Q.EOF then
  begin
    AxeJal := Q.Fields[0].AsString;
    if AxeJal <> Y_Axe.Value then
    begin
      Y_Journal.Value := '';
      AxeJal := '';
      HMCle.Execute(15, '', '');
    end else
      if Q.Fields[1].AsString = '' then
    begin
      Y_Journal.Value := '';
      AxeJal := '';
      HMCle.Execute(19, '', '');
    end;
  end;
  Ferme(Q);
end;

procedure TFVideSect.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;

procedure TFVideSect.Y_DATECOMPTABLEKeyPress(Sender: TObject; var Key: Char);
begin
  ParamDate(Self, Sender, Key);
end;

procedure TFVideSect.S_CLEREPARTITIONChange(Sender: TObject);
var
  Q: TQuery;
begin
  if S_CLEREPARTITION.Value = '' then
  begin
    GENERECEPT.Text := '';
    Exit;
  end;
  Q := OpenSQL('SELECT RE_COMPTES FROM CLEREPAR WHERE RE_AXE="' + Y_AXE.Value + '" AND RE_CLE="' + S_CLEREPARTITION.Value + '"', True);
  if not Q.EOF then GENERECEPT.Text := Q.Fields[0].AsString;
  Ferme(Q);
end;



end.

