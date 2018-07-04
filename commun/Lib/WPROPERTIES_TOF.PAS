{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 04/07/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : WPROPERTIES ()
Mots clefs ... : TOF;WPROPERTIES
*****************************************************************}
Unit WPROPERTIES_TOF ;

Interface

Uses
  StdCtrls,
  Controls,
  Classes,
  {$IFNDEF EAGLCLIENT}
    db,
    {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
    EdtREtat,
    Fe_Main,
  {$ELSE}
    UtilEagl,
    MainEAgl,
  {$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  uTob,
  wTof,
  HTB97,
  Dialogs,
  Menus
  ;

Type
  TOF_WPROPERTIES = Class (tWTOF)
    procedure OnNew                ; override;
    procedure OnDelete             ; override;
    procedure OnUpdate             ; override;
    procedure OnLoad               ; override;
    procedure OnArgument(S: String); override;
    procedure OnClose              ; override;
  private
    Prefixe  	 : string;
    Identifiant: string;
    Clef       : string;
   	TableName	 : string;
    CreeLe   	 : string;
    CreePar  	 : string;
    ModifieLe	 : string;
    ModifiePar : string;
    DBListeTableName: String;
    TobField	 : Tob;
    TobWPT		 : Tob;
    isDecomposed: Boolean;
    FormCall : TForm;
    FSearch: TFindDialog;
    FirstFind: Boolean;

    {tob}
    procedure InitTobs;
    procedure KillTobs;
    function wGetTobWPT(TarataTob: Tob): Boolean;

		{ Impression }
    procedure mnEditProp_OnClick(Sender: tObject);
    procedure mnListProp_OnClick(Sender: tObject);

    procedure LBMulDBListe_OnClick(Sender: TObject);
    procedure LBTABLENAME_OnClick(Sender: TObject);

    { Evénements }
    procedure BtSearch_OnClick(Sender: TObject);
    procedure Form_OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FSearch_Find(Sender: TObject);
    procedure PageControl_OnChange(Sender: TObject);
	end;

procedure wGetTobField(TableName: string; TobField: tob);

Implementation

uses
	wCommuns,
  {$IFDEF EAGLCLIENT}
  eMul,
  eFiche,
  eFichList,
  {$ELSE}
  Mul,
  Fiche,
  FichList,
  {$ENDIF}
  SaisieList,
  Vierge,
  Windows,
  graphics,
  UtilPGI
  ,CbpMCD
  ,CbpEnumerator
;

procedure TOF_WPROPERTIES.OnNew;
begin
  Inherited;
end;

procedure TOF_WPROPERTIES.OnDelete;
begin
  Inherited;
end;

procedure TOF_WPROPERTIES.OnUpdate;
begin
  Inherited;
end;

procedure TOF_WPROPERTIES.OnLoad;
var
  Q: TQuery;
  TS: TStringList;
begin
  Inherited;

  { Permet d'afficher les infos relatifs à la fiche dans l'onglet SAV }
  if V_PGI.SAV and Assigned(FormCall) then
  begin
    SetControlCaption('LBFICHENAME', FormCall.Name);
    if FormCall is TFVierge then
    begin
      SetControlVisible('PNVIE', True);
      SetControlCaption('LBTYPE', 'Fiche vierge');
      SetControlCaption('LBVIETOF', TFVierge(FormCall).TOFName);
    end
    else if (FormCall is TFFiche) or (FormCall is TFFicheListe) then
    begin
      SetControlVisible('PNFIC', True);
      SetControlCaption('LBTYPE', iif(FormCall is TFFiche, 'Fiche principale liée à une table', 'Fiche liste'));
      if FormCall is TFFiche then
        SetControlCaption('LBFICCLE', TFFiche(FormCall).UniqueName)
      else
        SetControlCaption('LBFICCLE', TFFicheListe(FormCall).UniqueName);
    end
    else if FormCall is TFMul then
    begin
      DBListeTableName := '';
      Q := OpenSQL('SELECT LI_DATA FROM LISTE WHERE LI_LISTE="' + TFMul(FormCall).DBListe + '"', True, 1);
      try
        if not Q.Eof then
        begin
          TS := TStringList.Create;
          try
            TS.SetText(PChar(Q.FindField('LI_DATA').AsString));
            if TS.Count > 0 then
              DBListeTableName := Trim(TS.Strings[0]);
          finally
            TS.Free;
          end;
        end;
      finally
        Ferme(Q);
      end;

      if (VueToNum(DBListeTableName) <> -1) and Assigned(GetControl('LBMULDBLISTE')) then
      begin
        THLabel(GetControl('LBMULDBLISTE')).OnClick := LBMulDBListe_OnClick;
        THLabel(GetControl('LBMULDBLISTE')).Cursor := crHandPoint;
        THLabel(GetControl('LBMULDBLISTE')).Font.Color := clBlue;
        THLabel(GetControl('LBMULDBLISTE')).Font.Style := [fsUnderline];
      end;

      SetControlVisible('PNMUL', True);
      SetControlCaption('LBTYPE', 'Fiche multicritères');
      SetControlCaption('LBMULDBLISTE', TFMul(FormCall).DBListe + iif(VueToNum(DBListeTableName) <> -1, ' (' + TraduireMemoire('Vue') + ')', ''));
      SetControlCaption('LBMULTOF', TFMul(FormCall).TOFName);
    end
    else if FormCall is TFSaisieList then
    begin
      SetControlVisible('PNFSL', True);
      SetControlCaption('LBTYPE', 'Fiche de saisie en liste');
      SetControlCaption('LBFSLTABLETET', TFSaisieList(FormCall).TableEntete);
      SetControlCaption('LBFSLTABLELINK', TFSaisieList(FormCall).TableLink);
      SetControlCaption('LBFSLTABLENAME', TFSaisieList(FormCall).TableName);
      SetControlCaption('LBFSLTOF', TFSaisieList(FormCall).TOFName);
      SetControlCaption('LBFSLTREELISTE', TFSaisieList(FormCall).TreeListe);
      SetControlCaption('LBFSLCLE', TFSaisieList(FormCall).UniqueName);
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 13/09/2002
Modifié le ... :   /  /
Description .. : Gestion de la TobWPT
Mots clefs ... :
*****************************************************************}
procedure TOF_WPROPERTIES.InitTobs;
begin
	TobField := Tob.Create('FIELDS', nil, -1);
	TobWPT   := Tob.Create('WPT', nil, -1);
end;

procedure TOF_WPROPERTIES.KillTobs;
begin
  if Assigned(TobField) then TobField.Free;
	if Assigned(TobWPT) then TobWPT.Free;
end;

function TOF_WPROPERTIES.wGetTobWPT(TarataTob: Tob): Boolean;
var
	i: Integer;
  T: Tob;
begin
	for i := 0 to TobField.Detail.Count - 1 do
  begin
    T := Tob.Create('_Propriétés_', TarataTob, -1);
    if CreeLe = '' then
    	CreeLe := DateToStr(iDate1900);
    if ModifieLe = '' then
    	ModifieLe := DateToStr(iDate1900);
	  T.AddChampSupValeur('WPT_UTILISATEUR', V_PGI.User);
    T.AddChampSupValeur('WPT_ORDRE', i);
    T.AddChampSupValeur('WPT_TABLE', TableName);
    T.AddChampSupValeur('WPT_PREFIXE', Prefixe);
    T.AddChampSupValeur('WPT_CLEF', Identifiant);
    T.AddChampSupValeur('WPT_DATEC', StrToDate(CreeLe));
    T.AddChampSupValeur('WPT_CREEPAR', CreePar);
    T.AddChampSupValeur('WPT_DATEM', StrToDate(ModifieLe));
    T.AddChampSupValeur('WPT_MODIFPAR', ModifiePar);
    T.AddChampSupValeur('WPT_CHAMP', TobField.Detail[i].GetString('CHAMP'));
    T.AddChampSupValeur('WPT_LIBELLE', TobField.Detail[i].GetString('LIBELLE'));
    T.AddChampSupValeur('WPT_TYPE', TobField.Detail[i].GetString('TYPE'));
    T.AddChampSupValeur('WPT_VALEUR', TobField.Detail[i].GetString('VALEUR'));
  end;

  Result := TobField.Detail.Count = TobWPT.Detail.Count;
end;

procedure TOF_WPROPERTIES.OnArgument(S: String);
var
  sPtrFiche: String;

  procedure GetInfo;
  var
		Sql: string;
    TobInfo: Tob;
  begin
    Sql := 'SELECT ' + Prefixe + '_CREATEUR, ' + Prefixe + '_UTILISATEUR, ' + Prefixe + '_DATECREATION, ' + Prefixe + '_DATEMODIF'
         + ' FROM ' + TableName
         + ' WHERE ' + wMakeWhereSQL(Clef, Identifiant)
         ;
    TobInfo := Tob.Create('INFO', nil, -1);
    try
      if TobInfo.LoadDetailDBFromSql('INFO', Sql) then
      with TobInfo.Detail[0] do
      begin
        if GetString(Prefixe + '_CREATEUR') <> '' then
          CreePar := RechDom('TTUTILISATEUR', GetString(Prefixe + '_CREATEUR'), False)
        else
          CreePar := TraduireMemoire('Inconnu');
        CreeLe := FormatDateTime('dd/mm/yyyy', GetDateTime(Prefixe+'_DATECREATION'));

        { pour gérer les fiches jamais modifiées (code utilisateur = '000') }
        if (GetString(Prefixe + '_UTILISATEUR') <> '') and (GetString(Prefixe + '_UTILISATEUR') <> '000') then
          ModifiePar := RechDom('TTUTILISATEUR', GetString(Prefixe + '_UTILISATEUR'), False)
        else
          ModifiePar := TraduireMemoire('Inconnu');
        ModifieLe := FormatDateTime('dd/mm/yyyy', GetDateTime(Prefixe+'_DATEMODIF'));
      end;
    finally
      TobInfo.Free;
    end;
  end;

  procedure wGetValueTobField;
  var
  	Sql: string;
    Q	: tQuery;
    T  : Tob;
  	i		: integer;
  begin
  	Sql := 'SELECT *'
         + ' FROM ' + TableName
         + ' WHERE ' + wMakeWhereSQL(Clef, Identifiant)
         ;
		T := Tob.Create(TableName, nil, -1);
    Q := OpenSql(Sql, True);
    try
      if not Q.Eof then
      begin
        T.SelectDB('', Q);
        for i := 0 to tobField.Detail.Count - 1 do
          tobField.Detail[i].AddChampSupValeur('VALEUR', T.GetValue(tobField.Detail[i].GetValue('CHAMP')));
      end;
    finally
    	Ferme(Q);
      T.Free;
    end;
  end;

begin
  Inherited ;

  isDecomposed := False;

  {événement}
  if Assigned(GetControl('MNEDITPROP')) then
    tMenuItem(GetControl('MNEDITPROP')).OnClick := mnEditProp_OnClick;
  if Assigned(GetControl('MNLISTPROP')) then
    tMenuItem(GetControl('MNLISTPROP')).OnClick := mnListProp_OnClick;

  Prefixe     := GetArgumentValue(S, 'PREFIXE');
  Clef        := GetArgumentValue(S, 'CLEF');
  Clef        := StringReplace(Clef, '~', ';', [rfIgnoreCase, rfReplaceAll]);
  Identifiant := GetArgumentValue(S, 'IDENTIFIANT');
  Identifiant := StringReplace(Identifiant, '~', ';', [rfIgnoreCase, rfReplaceAll]);

  sPtrFiche := GetArgumentValue(S, 'PTRFICHE');
  if V_PGI.SAV and (sPtrFiche <> ''){and Assigned(VH_GP.TobTemp)} then
    FormCall := TForm(StrToInt(sPtrFiche))
  else
    FormCall := nil;

  TTabSheet(GetControl('TSSav')).TabVisible := Assigned(FormCall);

  TableName := PrefixeToTable(Prefixe);

  GetInfo;

  { Put ecran }
  if Assigned(GetControl('LBTABLENAME')) then
  begin
    ThLabel(GetControl('LBTABLENAME')).Caption := TableName;
    if V_PGI.SAV then
    begin
      ThLabel(GetControl('LBTABLENAME')).OnClick := LBTABLENAME_OnClick;
      THLabel(GetControl('LBTABLENAME')).Cursor := crHandPoint;
      THLabel(GetControl('LBTABLENAME')).Font.Color := clBlue;
      THLabel(GetControl('LBTABLENAME')).Font.Style := [fsUnderline];
    end;
  end;
  if Assigned(GetControl('LBPREFIXE')) then
    ThLabel(GetControl('LBPREFIXE')).Caption := Prefixe;
  if Assigned(GetControl('LBCLEF')) then
    ThLabel(GetControl('LBCLEF')).Caption := Identifiant;
  if Assigned(GetControl('LBCREELE')) then
    ThLabel(GetControl('LBCREELE')).Caption := CreeLe;
  if Assigned(GetControl('LBCREEPAR')) then
    ThLabel(GetControl('LBCREEPAR')).Caption := CreePar;
  if Assigned(GetControl('LBMODIFIELE')) then
    ThLabel(GetControl('LBMODIFIELE')).Caption := ModifieLe;
  if Assigned(GetControl('LBMODIFIEPAR')) then
    ThLabel(GetControl('LBMODIFIEPAR')).Caption := ModifiePar;
  if Assigned(GetControl('BTSEARCH')) then
    TToolBarButton97(GetControl('BTSEARCH')).OnClick := BtSearch_OnClick;
  if Assigned(GetControl('PCPROPERTIES')) then
    TPageControl(GetControl('PCPROPERTIES')).OnChange := PageControl_OnChange;
  Ecran.OnKeyDown := Form_OnKeyDown;

  FSearch := TFindDialog.Create(Ecran);
  FSearch.OnFind := FSearch_Find;

  { Grid }
  InitTobs;
  wGetTobField(TableName, tobField);
  wGetValueTobField;
  wGetTobWPT(TobWPT);
	TobWPT.PutGridDetail(ThGrid(GetControl('GRID')), false, false, 'WPT_CHAMP;WPT_VALEUR;WPT_LIBELLE;WPT_TYPE');
end ;

procedure TOF_WPROPERTIES.OnClose ;
begin
  Inherited ;
  KillTobs;
end ;

procedure TOF_WPROPERTIES.mnEditProp_OnClick(Sender: tObject);
begin
	LanceEtatTob('E', 'WD0', 'PT0', TobWPT, True, False, False, nil, '', TraduireMemoire('Propriétés'), False);
end;

procedure TOF_WPROPERTIES.mnListProp_OnClick(Sender: tObject);
begin
	LanceEtatTob('E', 'WD0', 'PT1', TobWPT, True, False, False, nil, '', TraduireMemoire('Propriétés'), False);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 04/07/2002
Modifié le ... :   /  /
Description .. : Charge la Tob avec les noms de champs
Mots clefs ... :
*****************************************************************}
procedure wGetTobField(TableName: string; TobField: tob);
var
	i: integer;
  t: tob;
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
  FieldList : IEnumerator ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  //
  table := Mcd.getTable(TableName);
  FieldList := Table.Fields;
  FieldList.Reset();
  While FieldList.MoveNext do
  begin
  	if (FieldList.Current as IFieldCOM).name <> '' then
    begin
	   	t := Tob.Create('FIELD', TobField, -1);
  	  t.AddChampSupValeur('CHAMP', (FieldList.Current as IFieldCOM).name);
      t.AddChampSupValeur('LIBELLE', (FieldList.Current as IFieldCOM).Libelle);
      t.AddChampSupValeur('TYPE', (FieldList.Current as IFieldCOM).Tipe);
    end;
	end;
end;

procedure TOF_WPROPERTIES.LBMulDBListe_OnClick(Sender: TObject);
begin
  if DBListeTableName <> '' then
    AGLLanceFiche('W','WBLOCNOTE', '', '', 'ACTION=' + iif(Action = '', 'MODIFICATION', Action) + ';RO=' + wTrue + ';PROPFONT=' + wTrue + ';LASTRING=' + V_PGI.DEVues[VueToNum(DBListeTableName)].SQL);
end;

procedure TOF_WPROPERTIES.LBTABLENAME_OnClick(Sender: TObject);
begin
  AGLLanceFiche('W','WBLOCNOTE', '', '', 'ACTION=' + iif(Action = '', 'MODIFICATION', Action) + ';RO=' + wTrue + ';PROPFONT=' + wTrue + ';LASTRING=' + wMakeFieldString(THLabel(Sender).Caption, #13 + #10));
end;

procedure TOF_WPROPERTIES.BtSearch_OnClick(Sender: TObject);
begin
  FSearch.Execute;
  FirstFind := True;
end;

procedure TOF_WPROPERTIES.Form_OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if GetControlEnabled('BTSEARCH') then
  begin
    if (Key = vk_Cherche) and (ssCtrl	in Shift) then
    begin
      FSearch.Execute;
      FirstFind := True;
    end
    else if (Key = VK_F3) then
      FSearch_Find(FSearch)
    else
      TFVierge(Ecran).FormKeyDown(Sender, Key, Shift)
  end
end;

procedure TOF_WPROPERTIES.FSearch_Find(Sender: TObject);
begin
  Rechercher(THGrid(GetControl('GRID')), FSearch, FirstFind);
end;

procedure TOF_WPROPERTIES.PageControl_OnChange(Sender: TObject);
begin
  SetControlEnabled('BTSEARCH', TPageControl(Sender).ActivePage.Name = 'TSDONNEES');
  if not GetControlEnabled('BTSEARCH') then
    FSearch.CloseDialog()
end;

Initialization
  registerclasses([TOF_WPROPERTIES]);
end.

