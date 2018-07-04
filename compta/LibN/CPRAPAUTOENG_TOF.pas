{***********UNITE*************************************************
Auteur  ...... : Thong Hor LIM
Créé le ...... : 16/05/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPRAPAUTOENG ()
Mots clefs ... : TOF;CPRAPAUTOENG
*****************************************************************}
Unit CPRAPAUTOENG_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fe_Main,     // AGLLanceFiche
     HDB,       // THDBGrid
{$else}
     eMul,
     MaineAGL,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     Ent1,            // VH^, TFichierBase
     Ed_Tools,        // Pour le videListe
     Htb97,           // TToolBarButton97, TToolWindow97
     Grids,
     HQry,
     HMsgBox,
     UTOF,
     UTob,
     LookUp,          // LookUpList
     HCompte,
     CPGESTENG_TOF,
     SaisComm,        // TobToIdent, MvtToIdent
     SAISUTIL;        // RMVT

Type
  TOF_CPRAPAUTOENG = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  private
    FFirstDisplay            : Boolean;
    FOnPlanComptaChange      : TNotifyEvent;
    FOnMontantKeyPress       : TKeyPressEvent;
    FOnSeuilKeyPress         : TKeyPressEvent;

  {$IFDEF EAGLCLIENT}
    FListe : THGrid;
  {$ELSE}
    FListe : THDBGrid;
  {$ENDIF EAGLCLIENT}
    BChercher          :TToolbarButton97;

    //onglet comptes
    PlanCompta         :THValCombobox;
    AUXILIAIRE         :THEdit;
    AUXILIAIRE_        :THEdit;

    XX_WHERE           :THEdit ;

    //onglet critères
    CBMntSeuil,
    CBRefCmd,
    CBAutreCrit        :TCheckBox;
    LeSeuil,
    LeMontant          :THEdit;
    CB_Criteres        :array[0..4] of THValCombobox;
    //onglet autres critères
    CB_Tables          :array[0..3] of TCheckBox;

    function  InitControles: boolean ;
    procedure InitComboAutreCrit(Combo :THValCombobox; QLstInfo: TQuery);
    function  CritereMontantToSQL(PrefixEcr, PrefixEng, EngValue: String): String;
    function  CritereSeuilToSQL(PrefixEcr, PrefixEng, EngValue: String): String;
    function  CritereRefCmdToSQL(PrefixEcr, EngValue: String): String;
    function  CritereAutreCritereToSQL(Indice: Integer; PrefixEcr, EngValue: String): String;
    function  CritereTableLibreToSQL(Indice: Integer; PrefixEcr, EngValue: String): String;
    function  CritereStandardToSQL(PrefixEcr, PrefixEng, Auxiliaire, General, Devise: String): String;
    procedure LanceRapproManuel;


    procedure OnPlanComptaChange(Sender: TObject);
    procedure OnCBMntSeuilClick(Sender: TObject);
    procedure OnCBAutreCritClick(Sender: TObject);
    procedure OnOuvrirClick(Sender: TObject);
    procedure OnMontantSeuilKeyPress(Sender: TObject; var Key: Char);
  end ;


procedure CPLanceFiche_CPRAPAUTOENG;

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}
  utilPGI; // OpenSelect

procedure CPLanceFiche_CPRAPAUTOENG;
begin
  AGLLanceFiche('CP', 'CPRAPPAUTOENG_MUL', '', '', '') ;
end;

{Voir unité UTILEDT}
procedure PositionneFourchetteST(TC1,TC2 : THEdit; tt: TZoomTable) ;
var
  St:          String;
  Q:           TQuery;
begin
  if (TC1.Text='') And (TC2.Text='') then begin
    Case CaseFic(tt) Of
      fbGene : St:='SELECT MIN(G_GENERAL), Max(G_GENERAL) FROM GENERAUX WHERE G_FERME="-" ' ;
      fbAux : St:='SELECT MIN(T_AUXILIAIRE), Max(T_AUXILIAIRE) FROM TIERS WHERE T_FERME="-" ' ;
      fbJal : St:='SELECT MIN(J_JOURNAL), Max(J_JOURNAL) FROM JOURNAL WHERE J_FERME="-" ' ;
      fbAxe1..fbAxe5 : St:='SELECT MIN(S_SECTION), Max(S_SECTION) FROM SECTION WHERE S_FERME="-" ' ;
      fbBudGen : St:='SELECT MIN(BG_BUDGENE), Max(BG_BUDGENE) FROM BUDGENE WHERE BG_FERME="-" ' ;
      fbBudJal : St:='SELECT MIN(BJ_BUDJAL), Max(BJ_BUDJAL) FROM BUDJAL WHERE BJ_FERME="-" ' ;
      fbBudSec1..fbBudSec5 : St:='SELECT MIN(BS_BUDSECT), Max(BS_BUDSECT) FROM BUDSECT WHERE BS_FERME="-" ' ;
      fbNatCpt : St:='SELECT MIN(NT_NATURE), Max(NT_NATURE) FROM NATCPTE WHERE NT_SOMMEIL="-" ' ;
      end;
    St:=St+RecupWhere(tt) ;

    {**************************************************************************}
    {Il faut modifier l'unité HCompte, la fonction RecupWhere car la nature
    d'un jnl de vente est VTE et pas VEN}
    {**************************************************************************}
    if tt = tzJvente then
      St := StringReplace(St, '"VEN"', '"VTE"', []);
    {**************************************************************************}
    {**************************************************************************}
    Q:=OpenSQL(St,TRUE) ;
    if not Q.EOF then begin
      TC1.Text := Q.Fields[0].AsString;
      TC2.Text := Q.Fields[1].AsString;
      end;
    Ferme(Q);
    end;
end;

{Provient de LettUtil}
procedure InitTablesLibresEcriture( TT : TTabSheet ) ;
Var LesLib : HTStringList ;
    i : Integer ;
    St,Titre,Ena : String ;
    Trouver : Boolean ;
    LL      : TCheckBox ;
BEGIN
if TT=Nil then Exit ; Trouver:=False ;
LesLib:=HTStringList.Create ;
GetLibelleTableLibre('E',LesLib) ;
for i:=0 to LesLib.Count-1 do
    BEGIN
    St:=LesLib.Strings[i] ; Titre:=ReadTokenSt(St) ; Ena:=St ;
    LL:=TCheckBox(TForm(TT.Owner).FindComponent('CBTABLE'+IntToStr(i+1))) ;
    if LL<>Nil then
       BEGIN
       LL.Caption:=Titre ; LL.Enabled:=(Ena='X') ;
       if ((EstSerie(S3)) and (i>2)) then
         LL.Visible:=False
       else
         Trouver := True;
       END ;
    END ;
TT.TabVisible:=Trouver ;
LesLib.Clear ; LesLib.Free ;
END ;

procedure TOF_CPRAPAUTOENG.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPRAPAUTOENG.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPRAPAUTOENG.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPRAPAUTOENG.OnLoad ;
var
  CritereSQL:    String;
  i:             Integer;
begin
  Inherited ;

  if FFirstDisplay then       {A la première ouverture du formulaire, la grille doit rester vide}
    begin
    XX_WHERE.Text := '(1=2)';
    FFirstDisplay := False;
    Exit;
    end;

  if PlanCompta.ItemIndex = 0 then
    PositionneFourchetteST(AUXILIAIRE, AUXILIAIRE_, tzTfourn)
  else if PlanCompta.ItemIndex = 1 then
    PositionneFourchetteST(AUXILIAIRE, AUXILIAIRE_, tzTclient);

  CritereSQL := CritereMontantToSQL('E1.', 'CENGAGEMENT.', 'CPECRLIEENG.E_DEBITDEV-CPECRLIEENG.E_CREDITDEV')+
                CritereSeuilToSQL(  'E1.', 'CENGAGEMENT.', 'abs(CPECRLIEENG.E_DEBITDEV-CPECRLIEENG.E_CREDITDEV)')+
                CritereRefCmdToSQL( 'E1.', 'CPECRLIEENG.E_REFINTERNE');
  for i:=0 to High(CB_Criteres) do
    CritereSQL := CritereSQL + CritereAutreCritereToSQL(i, 'E1.', 'CPECRLIEENG.'+CB_Criteres[i].Value);

  for i:=0 to High(CB_Tables) do
    CritereSQL := CritereSQL + CritereTableLibreToSQL(i, 'E1.', 'CPECRLIEENG.E_TABLE'+IntToStr(i));

  XX_WHERE.Text := 'E_QUALIFPIECE="P"'+
    ' and (CEN_STATUTENG="E" or CEN_STATUTENG="P")'+
    ' and exists(select E1.E_EXERCICE '+
          ' from ECRITURE E1 left outer join CENGAGEMENT on E1.E_EXERCICE = CENGAGEMENT.CEN_EXERCICE' +
                   ' and E1.E_JOURNAL = CENGAGEMENT.CEN_JOURNAL' +
                   ' and E1.E_DATECOMPTABLE = CENGAGEMENT.CEN_DATECOMPTABLE' +
                   ' and E1.E_NUMEROPIECE = CENGAGEMENT.CEN_NUMEROPIECE' +
                   ' and E1.E_NUMLIGNE = CENGAGEMENT.CEN_NUMLIGNE' +
                   ' and E1.E_QUALIFPIECE = CENGAGEMENT.CEN_QUALIFPIECE' +
                   ' and E1.E_NUMECHE = CENGAGEMENT.CEN_NUMECHE'+
          ' where '+
           CritereStandardToSQL('E1.', 'CENGAGEMENT.', 'CPECRLIEENG.E_AUXILIAIRE',
              'CPECRLIEENG.E_GENERAL', 'CPECRLIEENG.E_DEVISE') +
           CritereSQL+')';

end ;

procedure TOF_CPRAPAUTOENG.OnArgument (S : String ) ;
var
  i: Integer;
begin
  Inherited ;

  //Elements de la  TOF.
{$IFDEF EAGLCLIENT}
  FListe := THGrid(GetControl('FListe', True));
{$ELSE}
  FListe := THDBGrid( GetControl('FListe', True));
{$ENDIF EAGLCLIENT}
  BChercher     := TToolbarButton97(GetControl('BCHERCHE', True));
  PlanCompta    := THValComboBox(GetControl('PlanCompta', True));
  AUXILIAIRE    := THEdit(GetControl('E_AUXILIAIRE', True));
  AUXILIAIRE_   := THEdit(GetControl('E_AUXILIAIRE_', True));
  CBMntSeuil    := TCheckBox(GetControl('CBMontant', True));
  CBRefCmd      := TCheckBox(GetControl('CBRefCde', True));
  CBAutreCrit   := TCheckBox(GetControl('CBAutreCrit', True));
  LeSeuil       := THEdit(GetControl('LeSeuil', True));
  LeMontant     := THEdit(GetControl('LeMontant', True));
  XX_WHERE      := THEdit(GetControl('XX_WHERE', True));

  for i:=0 to High(CB_Criteres) do
    CB_Criteres[i] := THValComboBox(GetControl('CRITERE'+IntToStr(i+1), True));

  for i:=0 to High(CB_Tables) do
    CB_Tables[i] := TCheckBox(GetControl('CBTABLE'+IntToStr(i+1), True));

  PlanCompta.ItemIndex := 0;    {Fourn/Client}
  FOnPlanComptaChange  := PlanCompta.OnChange;
  PlanCompta.OnChange  := OnPlanComptaChange;

  FOnMontantKeyPress := LeMontant.OnKeyPress;
  FOnSeuilKeyPress   := LeSeuil.OnKeyPress;
  LeMontant.OnKeyPress := OnMontantSeuilKeyPress;
  LeSeuil.OnKeyPress   := OnMontantSeuilKeyPress;

  CBMntSeuil.OnClick   := OnCBMntSeuilClick;
  CBAutreCrit.OnClick := OnCBAutreCritClick;

  TToolbarButton97(GetControl('BOUVRIR')).OnClick := OnOuvrirClick;
  FListe.OnDblClick := OnOuvrirClick;

  SetControlText('E_Devise', V_PGI.DevisePivot);
  FFirstDisplay := True;

  TFMul(Ecran).HelpContext := 0 ;
  InitControles;
end ;

procedure TOF_CPRAPAUTOENG.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CPRAPAUTOENG.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPRAPAUTOENG.OnCancel () ;
begin
  Inherited ;
end ;

function TOF_CPRAPAUTOENG.InitControles: boolean;
var
  Q:    TQuery ;
  i:    Integer;
begin
  Result := True;

  InitTablesLibresEcriture(GetControl('PTABLIBRE') as TTabSheet);

  Q := OpenSQL(
         'select CO_CODE, CO_ABREGE, DH_LIBELLE' +
         ' from COMMUN join DECHAMPS on CO_ABREGE = DH_NOMCHAMP'+
         ' where CO_TYPE="CRE"'+
         ' order by CO_CODE', True);
  try
    for i:=0 to High(CB_Criteres) do
      InitComboAutreCrit(CB_Criteres[i], Q);
  finally
    ferme(Q);
    end;
end;

procedure TOF_CPRAPAUTOENG.InitComboAutreCrit(Combo :THValCombobox; QLstInfo: TQuery);
var
  i:    integer;
begin
  QLstInfo.First;
  while not QLstInfo.Eof do
    begin
    for i:=0 to Combo.Items.Count-1 do
      begin
      if QLstInfo.FindField('CO_CODE').AsString = Combo.Values[i] then
        begin
        Combo.Values[i] := QLstInfo.FindField('CO_ABREGE').AsString;
        Combo.Items[i]  := QLstInfo.FindField('DH_LIBELLE').AsString;
        break;
        end;
      end;
    QLstInfo.Next;
    end;
end;

procedure TOF_CPRAPAUTOENG.OnPlanComptaChange(Sender: TObject);
var
  Tablette: String;
begin
  if PlanCompta.ItemIndex = 0 then
    Tablette  := 'TZTFOURN'
  else if PlanCompta.ItemIndex = 1 then
    Tablette  := 'TZTCLIENT';

  AUXILIAIRE.Text      := '';
  AUXILIAIRE_.Text     := '';
  AUXILIAIRE.DataType  := Tablette;
  AUXILIAIRE_.DataType := Tablette;

  if Assigned(FOnPlanComptaChange) then
    FOnPlanComptaChange(Sender);
end;

procedure TOF_CPRAPAUTOENG.OnCBAutreCritClick(Sender: TObject);
var
  Visible: Boolean;
  i:       Integer;
begin
  Visible := (Sender as TCheckBox).Checked;
  for i:=0 to High(CB_Criteres) do
    begin
    CB_CRITERES[i].Visible := Visible;
    if Visible then
      CB_CRITERES[i].ItemIndex := -1;
    end;
end;

procedure TOF_CPRAPAUTOENG.OnCBMntSeuilClick(Sender: TObject);
var
  Visible: Boolean;
begin
  Visible := (Sender as TCheckBox).Checked;
  SetControlVisible('LSEUIL_PC', Visible);
  SetControlVisible('LSEUILMONTANT', Visible);
  LeSeuil.Visible := Visible;
  LeMontant.Visible := Visible;
end;

procedure TOF_CPRAPAUTOENG.OnMontantSeuilKeyPress(Sender: TObject; var Key: Char);
begin
  if Sender = LeSeuil then
    begin
    LeMontant.Text := '';
    if Assigned(FOnMontantKeyPress) then
      FOnMontantKeyPress(Sender, Key);
    end
  else
    begin
    LeSeuil.Text := '';
    if Assigned(FOnSeuilKeyPress) then
      FOnSeuilKeyPress(Sender, Key);
    end;
end;

function TOF_CPRAPAUTOENG.CritereMontantToSQL(PrefixEcr, PrefixEng, EngValue: String): String;
var
  SoldeFac:  String;
  SoldeEcr:  String;
  MntRappro: String;
begin
  Result := '';
  if (not CBMntSeuil.Checked) or (LeMontant.Text='') then
    Exit;

  {Facture partiellement rapprochée: abs((Eng debit-Eng crédit)-(Fact debit-Fact credit-Fact rappro)) <= Montant.
       Dans le cas d'un engagement HT, Fact rappro tient compte de la TVA.
   Facture non rapprochée: abs((Eng debit-Eng crédit)-(Fact debit-Fact credit)) <= Montant,
       Dans le cas d'un engagement HT, il faudrait connaitre la TVA à déduire en devise}

  MntRappro := PrefixEng+'CEN_MONTANTRAP';
  SoldeFac  := PrefixEcr+'E_DEBITDEV+'+PrefixEcr+'E_CREDITDEV+'+MntRappro;
  SoldeEcr  := PrefixEcr+'E_DEBITDEV+'+PrefixEcr+'E_CREDITDEV';
  Result :=
    ' and (    abs('+EngValue+'-'+SoldeFac+')<='+StrFPoint(StrToFloat(LeMontant.Text))+
          ' or ('+MntRappro+' is null and abs('+EngValue+'-'+SoldeEcr+')<='+StrFPoint(StrToFloat(LeMontant.Text))+')'+
          ' or ('+MntRappro+'=0 and abs('+EngValue+'-'+SoldeEcr+')<='+StrFPoint(StrToFloat(LeMontant.Text))+')'+
         ')';
end;

function TOF_CPRAPAUTOENG.CritereSeuilToSQL(PrefixEcr, PrefixEng, EngValue: String): String;
var
  SoldeFac:     String;
  SoldeEcr:     String;
  MntRappro:    String;
  PourcentMin:  Double;
  PourcentMax:  Double;
begin
  Result := '';
  if (not CBMntSeuil.Checked) or (LeSeuil.Text='') or (StrToFloat(LeSeuil.Text)=0) then
    Exit;

  MntRappro   := PrefixEng+'CEN_MONTANTRAP';
  SoldeFac    := PrefixEcr+'E_DEBITDEV-'+PrefixEcr+'E_CREDITDEV-'+MntRappro;
  SoldeEcr    := PrefixEcr+'E_DEBITDEV-'+PrefixEcr+'E_CREDITDEV';
  PourcentMin := 1-(StrToFloat(LeSeuil.Text)/100);
  PourcentMax := 1+(StrToFloat(LeSeuil.Text)/100);
  Result :=
      ' and (    abs('+SoldeFac+') between '+EngValue+'*'+StrFPoint(PourcentMin)+' and '+EngValue+'*'+StrFPoint(PourcentMax)+
            ' or ('+MntRappro+' is null and abs('+SoldeEcr+') between '+EngValue+'*'+StrFPoint(PourcentMin)+' and '+EngValue+'*'+StrFPoint(PourcentMax)+')'+
            ' or ('+MntRappro+'=0 and abs('+SoldeEcr+') between '+EngValue+'*'+StrFPoint(PourcentMin)+' and '+EngValue+'*'+StrFPoint(PourcentMax)+')'+
           ')';
end;

function TOF_CPRAPAUTOENG.CritereAutreCritereToSQL(Indice: Integer; PrefixEcr, EngValue: String): String;
begin
  Result := '';
  if not CBAutreCrit.Checked then
    Exit;

  if (CB_Criteres[Indice].ItemIndex<>-1) and (CB_Criteres[Indice].Text <> CB_Criteres[Indice].VideString) then
    Result := Result+' and '+PrefixEcr+CB_Criteres[Indice].Value+'='+EngValue;
end;

function TOF_CPRAPAUTOENG.CritereRefCmdToSQL(PrefixEcr, EngValue: String): String;
begin
  Result := '';
  if not CBRefCmd.Checked then
    Exit;

  Result := ' and '+PrefixEcr+'E_REFINTERNE='+EngValue;
end;

function TOF_CPRAPAUTOENG.CritereTableLibreToSQL(Indice: Integer; PrefixEcr, EngValue: String): String;
begin
  Result := '';
  if not CB_Tables[Indice].Checked then
    Exit;

  Result := ' and '+PrefixEcr+'E_TABLE'+IntToStr(Indice)+'='+EngValue;
end;

procedure TOF_CPRAPAUTOENG.OnOuvrirClick(Sender: TObject);
var
  Nbr        : Integer;
begin
  {$IFDEF EAGLCLIENT}
  Nbr := FListe.RowCount;
  {$ELSE}
  Nbr := FListe.DataSource.DataSet.RecordCount;
  {$ENDIF}
  if Nbr = 0 then
    begin
    PGIError(TraduireMemoire('Vous devez sélectionner un engagement.'));
    Exit;
    end;

  // TEST des valeurs de paramètres !
  // Critères sur montant
  if GetControlText('CBMONTANT')='X' then
    begin
    if Trim(GetControlText('LESEUIL'))='' then
      SetControlText('LESEUIL',   '0,0000' ) ;
    if Trim(GetControlText('LEMONTANT'))='' then
      SetControlText('LEMONTANT', '0,00' ) ;
    end ;
  // Autres Critères
  if GetControlText('CBAUTRECRIT')='X' then
    begin
    if not ( (GetControlText('CRITERE1')<>'') or (GetControlText('CRITERE2')<>'')
             or (GetControlText('CRITERE3')<>'') or (GetControlText('CRITERE4')<>'')
             or (GetControlText('CRITERE5')<>'') ) then
      begin
      PGIError(TraduireMemoire('Vous devez renseigner un critère de rapprochement .'));
      SetActiveTabSheet('PCOMPLEMENT') ;
      SetFocusControl('CRITERE1') ;
      Exit;
      end ;
    end ;

  LanceRapproManuel;
  BChercher.Click;
end;

function TOF_CPRAPAUTOENG.CritereStandardToSQL(PrefixEcr, PrefixEng, Auxiliaire, General, Devise: String): String;
var
  MntRappro:   String;
  SoldeFac:    String;
begin
  MntRappro   := PrefixEng+'CEN_MONTANTRAP';
  SoldeFac    := PrefixEcr+'E_DEBITDEV-'+PrefixEcr+'E_CREDITDEV-'+MntRappro;

  Result := PrefixEcr + 'E_QUALIFPIECE="N" AND '                // type
            + PrefixEcr + 'E_ECRANOUVEAU="N" AND '                  // Anouveau
            + '(' + PrefixEcr + 'E_NATUREPIECE="FC" OR '        // Natures 'AC;AF;FC;FF;OD;'
                  + PrefixEcr + 'E_NATUREPIECE="FF" OR '
                  + PrefixEcr + 'E_NATUREPIECE="OC" OR '
                  + PrefixEcr + 'E_NATUREPIECE="OD" OR '
                  + PrefixEcr + 'E_NATUREPIECE="OF" )'
            + ' AND ( ( ' + SoldeFac  + '<>0 )'                 // Facture non totalement rapprochée ou jamais rapprochée
              +  ' OR ( ' + MntRappro + ' IS NULL )'
              +  ' OR ( ' + MntRappro + '=0 ) ) '
            +  ' AND ' + PrefixEcr + 'E_AUXILIAIRE=' + Auxiliaire
            +  ' AND ' + PrefixEcr + 'E_GENERAL='    + General
            +  ' AND ' + PrefixEcr + 'E_DEVISE='     + Devise     ;

end;

procedure TOF_CPRAPAUTOENG.LanceRapproManuel;
var
  i:             Integer;
  TOBEng:        TOB;
  LTobEng:       TOB;
  QQ:            TQuery;
  Auxiliaire:    String;
  CritereEng:    String;
  CritereEcr:    String;
  Ok:            Boolean;
begin
  Ok := False;
  TOBEng:=TOB.Create('',Nil,-1);
  try
    QQ:=OpenSQL('SELECT * FROM ECRITURE'+
      ' WHERE E_JOURNAL="'       + GetField('E_JOURNAL')                   +'"'+
      '   AND E_EXERCICE="'      + GetField('E_EXERCICE')                  +'"'+
      '   AND E_DATECOMPTABLE="' + USDATETIME(GetField('E_DATECOMPTABLE')) +'"'+
      '   AND E_NUMEROPIECE='    + IntToStr(GetField('E_NUMEROPIECE'))     +
      '   AND E_NUMLIGNE='       + IntToStr(GetField('E_NUMLIGNE'))        +
      '   AND E_NUMECHE='        + IntToStr(GetField('E_NUMECHE'))         +
      '   AND E_QUALIFPIECE="'   + GetField('E_QUALIFPIECE')               +'"'
    ,True) ;
    TOBEng.LoadDetailDB('ECRITURE','','',QQ,False) ;
    Ferme(QQ) ;
    if TOBEng.Detail.Count>0 then
      begin
      LTobEng    := TOBEng.Detail[0];
      Auxiliaire := LTOBEng.GetValue('E_AUXILIAIRE');
      CritereEng := EncodeLC(TOBToIdent(LTobEng, True));
      CritereEcr :=
        CritereStandardToSQL('', '',
          '"'+LTOBEng.GetValue('E_AUXILIAIRE')+'"', '"'+LTOBEng.GetValue('E_GENERAL')+'"',
          '"'+LTOBEng.GetValue('E_DEVISE')+'"')+
        CritereMontantToSQL('', '',
//          StrFPoint(Abs(LTOBEng.GetValue('E_DEBITDEV')-LTOBEng.GetValue('E_CREDITDEV'))))+
          StrFPoint( LTOBEng.GetValue('E_DEBITDEV')-LTOBEng.GetValue('E_CREDITDEV') ) ) +
        CritereSeuilToSQL('', '',
          StrFPoint(Abs(LTOBEng.GetValue('E_DEBITDEV')-LTOBEng.GetValue('E_CREDITDEV'))))+
        CritereRefCmdToSQL('', '"'+LTOBEng.GetValue('E_REFINTERNE')+'"');
      for i:=0 to High(CB_Criteres) do
        CritereEcr := CritereEcr +
          CritereAutreCritereToSQL(i, '', '"'+LTOBEng.GetValue(CB_Criteres[i].Value)+'"');

      for i:=0 to High(CB_Tables) do
        CritereEcr := CritereEcr +
          CritereTableLibreToSQL(i, '', '"'+LTOBEng.GetValue('E_TABLE'+IntToStr(i))+'"');
      Ok := True;
      end;
  finally
    TOBEng.Free;
    end;

  if Ok then
    CPLanceFiche_CPGESTENG('RAPPROAUTO;'+Auxiliaire+';'+CritereEcr+'|'+CritereEng+'|');
end;

Initialization
  registerclasses ( [ TOF_CPRAPAUTOENG ] ) ;
end.

