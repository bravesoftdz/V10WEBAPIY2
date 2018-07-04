unit TofMulBudRub;

interface
uses
{$IFDEF EAGLCLIENT}
     MaineAGL,
     eMul,
     eTablette,
     CPAFFLISTRUB_TOF,     
{$ELSE}
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Mul,
     Fe_Main,
     Tablette,
     AffList,
{$ENDIF}
      Controls,StdCtrls,ComCtrls,forms,Graphics,Classes,sysutils,
      HCtrls,HEnt1,HMsgBox,UTOF,UTOB,LookUp,HDB,Vierge,HRichOLE, 
      UtilDiv,choix,
      HTB97;
Type
     TOF_ConvRubBud = class (TOF)
     private
       LaFamille: THedit ;
       Pages: TPageControl ;
  {$IFDEF EAGLCLIENT}
       FListe: THGrid ;
  {$ELSE}
       FListe: THDBGrid ;
  {$ENDIF}
       LSelected,LRubrique: TStringList ;
       ListeRub: Tob ;
       LMsg : THMsgBox ;
       BHelp: TToolbarButton97;
       procedure OnRubriqueClick(Sender : TObject) ;
       procedure OnBOuvrirClick(Sender : TObject) ;
       procedure ChargeLRubrique ;
       procedure OnBZoomClick(Sender : TObject) ;
       procedure OnElipsisClickLafamille(Sender : TObject) ;
       procedure ChargeLSelected ;
       procedure MiseAJourQuery(var Q2, Q1 : TOB) ;
       function ExisteRub(Rub : string) : boolean ;
       procedure InitMsg ;
       function AfficheMsg(num : integer;Av,Ap : string ) : Word ;
       procedure FermeMsg ;
       procedure BHelpClick(Sender: TObject);
     public
       procedure OnArgument  (stArgument : String ) ; override ;
       procedure OnLoad ; override ;
       procedure OnClose ; override ;
     end;

implementation

uses
    {$IFDEF eAGLCLIENT}
    MenuOLX, FamRub_Tof
    {$ELSE}
    MenuOLG,
      //{$IFDEF ESP}
      FAMRUB_TOF
      (*{$ELSE}
      FamRub
      {$ENDIF ESP}*)
    {$ENDIF eAGLCLIENT}
    ;

{---------------------------------------------------------------------------------------}
procedure TOF_ConvRubBud.OnArgument (stArgument : String ) ;
{---------------------------------------------------------------------------------------}
var
  CEdit : THEdit ;
  BButton : TButton ;
begin
  inherited ;
  TFMul(Ecran).HelpContext:=7812000;
  CEdit:=THedit(GetControl('RB_RUBRIQUE')) ;
  if (CEdit<>nil) and not assigned(CEdit.OnElipsisClick) then CEdit.OnElipsisClick:=OnRubriqueClick ;

  BButton:=TButton(GetControl('BOUVRIR')) ;
  if (BButton<>nil) then BButton.OnClick:=OnBOuvrirClick ;

  BButton:=TButton(GetControl('BZOOM')) ;
  if (BButton<>nil) and not assigned(BButton.OnClick) then BButton.OnClick:=OnBZoomClick ;

  LaFamille:=THEdit(GetControl('LAFAMILLE')) ;
  if (LaFamille<>nil) and not assigned(LaFamille.OnElipsisClick) then LaFamille.OnElipsisClick:=OnElipsisClickLafamille ;

  Pages:=TPageControl(GetControl('PAGES')) ;
  {$IFDEF EAGLCLIENT}
  FListe:=THGrid(GetControl('FLISTE')) ;
  {$ELSE}
  FListe:=THDBGrid(GetControl('FLISTE')) ;
  {$ENDIF}
  BHelp:=TToolbarButton97(GetControl('BAIDE')) ;
  if (BHelp<>nil) and (not Assigned(BHelp.OnClick)) then BHelp.OnClick:=BHelpClick ;

  LRubrique:=TStringList.Create ;
  LSelected:=TStringList.Create ;
  ListeRub:=Tob.Create('§LISTERUB',nil,-1) ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_ConvRubBud.OnLoad ;
{---------------------------------------------------------------------------------------}
Begin
  inherited ;
  InitMsg ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_ConvRubBud.OnClose ;
{---------------------------------------------------------------------------------------}
Begin
  inherited ;
  LRubrique.Free ;
  LSelected.Free ;
  ListeRub.Free ;
  FermeMsg ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_ConvRubBud.OnRubriqueClick(Sender : TObject) ;
{---------------------------------------------------------------------------------------}
begin
  LookUpList(TControl(Sender),'Liste des rubriques','RUBRIQUE','RB_RUBRIQUE','RB_LIBELLE','RB_NATRUB="BUD"','RB_RUBRIQUE',True,0) ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_ConvRubBud.OnElipsisClickLafamille(Sender : TObject) ;
{---------------------------------------------------------------------------------------}
var
  Fam, LibFam : String ;
begin
  Fam := LaFAmille.Text ;
  ParametrageFamilleRubrique('','',Fam,LibFam,False,TRUE) ;
  {if Fam <> '' Then} LaFAmille.Text := Fam ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_ConvRubBud.OnBOuvrirClick(Sender : TObject) ;
{---------------------------------------------------------------------------------------}
var
  i, NumInc    : Integer ;
  Q            : TQuery ;
  T1, TB1, TB2 : TOB ;
  Nom          : string ;
  Ret          : Word ;
  UpdateOk     : boolean;
Begin
  ChargeLRubrique ;
  ChargeLSelected ;
  NumInc:=0 ;

  if FListe.nbSelected < 1 then begin
    AfficheMsg(3,'','') ;
    Exit ;
  end;

  if LaFamille.Text = '' then begin
    AfficheMsg(0,'','') ;
    Pages.ActivePage := Pages.Pages[3];
    LaFamille.SetFocus ;
    Exit ;
  end ;

  Ret:=mrYes ;

  TB1 := TOB.Create('', nil, -1);
  TB2 := TOB.Create('RUBRIQUE', nil, -1);
  for i := 0 to LSelected.Count - 1 do begin
    Nom := 'GE:' + LSelected.Strings[i] ;

    {Si la rubrique n'existe pas, on fait un insert sinon Update}
    UpdateOk := ExisteRub(Nom);
    if UpdateOk and (Ret <> mrAll) then
      Ret := AfficheMsg(1, Nom, '');

    if Ret = mrNo then Continue ;

    Q := OpenSql('SELECT * FROM RUBRIQUE WHERE RB_RUBRIQUE="'+LSelected.Strings[i]+'"',False) ;
    TB1.LoadDetailDB('RUBRIQUE', '', '', Q, False);
    Ferme(Q);

    if TB1.Detail.Count > 0  then begin
      {On affecte les valeurs de TB1 à TB2 et MAJ de TB2}
      MiseAJourQuery(TB2, TB1);
      TB2.PutValue('RB_RUBRIQUE', Nom);
      TB2.PutValue('RB_FAMILLES', LaFamille.Text);

      if UpdateOk then TB2.UpdateDB
                  else TB2.InsertDb(nil);

      T1 := Tob.Create('$Fille',ListeRub,-1) ;
      T1.AddChampSup('RS_CODE',True) ;
      T1.PutValue('RS_CODE',TB1.Detail[0].GetValue('RB_RUBRIQUE')) ;
      T1.AddChampSup('RS_LIBELLE',True) ;
      T1.PutValue('RS_LIBELLE',TB1.Detail[0].GetValue('RB_LIBELLE')) ;
      T1.AddChampSup('RS_COMMENT',True) ;
      if UpdateOk then T1.PutValue('RS_COMMENT', 'Rubrique mise à jour')
                  else T1.PutValue('RS_COMMENT', 'Rubrique créée') ;
      Inc(NumInc) ;
    end ;
  end ;{for}

  FreeAndNil(TB1);
  FreeAndNil(TB2);
  Fliste.ClearSelected ;

  if NumInc <> 0 then begin
    AfficheListe(ListeRub,'Liste des rubriques integrées',['Nom Rubrique','Libellé']) ;
    AfficheMsg(2,'','') ;
    ListeRub.ClearDetail;
  end ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_ConvRubBud.OnBZoomClick(Sender : TObject) ;
{---------------------------------------------------------------------------------------}
begin
  {Appel de la tot CPRUBFAMILLE : UTOTRUBFAMILLE}
  ParamTable('CPRUBFAMILLE',taCreat,7775000,nil, 3,'Familles de rubriques') ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_ConvRubBud.ChargeLRubrique ;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery ;
  {$IFDEF EAGLCLIENT}
  i,
  n : Integer;
  T : TOB;
  {$ENDIF}
begin
  if LRubrique<>nil then begin
    LRubrique.Clear ;
    {$IFDEF EAGLCLIENT}
    T := Tob.Create('', nil, -1);
    with T do begin
      {On récupère toutes les rubriques}
      Q := OpenSql('SELECT RB_RUBRIQUE FROM RUBRIQUE ORDER BY RB_RUBRIQUE',True) ;
      if not Q.EOF then begin
        LoadDetailDb('', '', '', Q, False);
        n := Detail[0].GetNumChamp('RB_RUBRIQUE');
        {On boucle sur les rubriques pour charger la liste}
        for i := 0 to Detail.Count - 1 do
          LRubrique.Add(Detail[i].GetValeur(n));
      end;
      Ferme(Q) ;
    end;
    FreeAndNil(T);
   {$ELSE}
    Q:=OpenSql('SELECT RB_RUBRIQUE FROM RUBRIQUE ORDER BY RB_RUBRIQUE',True) ;
    while not Q.eof do begin
      LRubrique.Add(Q.Findfield('RB_RUBRIQUE').AsString) ;
      Q.Next ;
    end ;
    ferme(Q) ;
    {$ENDIF}
  end ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_ConvRubBud.ChargeLSelected ;
{---------------------------------------------------------------------------------------}
var
  i : integer ;
  {$IFDEF EAGLCLIENT}
  Fiche : TFMul;
  {$ELSE}
  Q : TQuery ;
  {$ENDIF}
begin
  {$IFDEF EAGLCLIENT}
  Fiche := TFMul(Ecran) ;
  if LSelected <> nil then begin
    LSelected.Clear ;
    for i:=0 to Fiche.FListe.nbSelected-1 do begin
      Fiche.FListe.GotoLeBookmark(i) ;
      Fiche.Q.TQ.Seek(FListe.Row - 1);
      LSelected.Add(Fiche.Q.Findfield('RB_RUBRIQUE').AsString) ;
    end ;
  end ;
  {$ELSE}
  if LSelected <> nil then begin
    Q := TQuery(FListe.DataSource.DataSet) ;
    LSelected.Clear ;
    for i:=0 to FListe.nbSelected-1 do
    begin
      FListe.GotoLeBookmark(i) ;
      LSelected.Add(Q.Findfield('RB_RUBRIQUE').AsString) ;
    end ;
  end ;
  {$ENDIF}
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_ConvRubBud.MiseAJourQuery(var Q2, Q1 : TOB) ;
{---------------------------------------------------------------------------------------}
var
  CptTmp,
  ExcTmp : string ;
  n : Integer;
begin
  for n := 0 to Q1.Detail[0].NbChamps - 1 do
    Q2.PutValeur(n, Q1.Detail[0].GetValeur(n));
    
  Q2.PutValue('RB_NATRUB', 'CPT');
  Q2.PutValue('RB_BUDJAL', '');
  Q2.PutValue('RB_PREDEFINI', 'DOS');
  Q2.PutValue('RB_NODOSSIER', '000000') ;

  Q2.PutValue('RB_NODOSSIER', V_PGI.NoDossier);

  if Q1.Detail[0].GetValue('RB_TYPERUB') = 'A/G' then begin
    Q2.PutValue('RB_TYPERUB', 'G/A');
    CptTmp := Q1.Detail[0].GetValue('RB_COMPTE1');
    ExcTmp := Q1.Detail[0].GetValue('RB_EXCLUSION1');
    Q2.PutValue('RB_COMPTE1', Q1.Detail[0].GetValue('RB_COMPTE2'));
    Q2.PutValue('RB_EXCLUSION1', Q1.Detail[0].GetValue('RB_EXCLUSION2'));
    Q2.PutValue('RB_COMPTE2', CptTmp);
    Q2.PutValue('RB_EXCLUSION2', ExcTmp);
  end ;
end ;

{---------------------------------------------------------------------------------------}
function TOF_ConvRubBud.ExisteRub(Rub : string) : boolean ;
{---------------------------------------------------------------------------------------}
var
  i : integer ;
begin
  Result:=False ;
  if LRubrique <> nil then begin
    for i := 0 to LRubrique.Count-1 do begin
      if Rub = LRubrique.Strings[i] then Result := True ;
      if (Result = True) or (Rub < LRubrique.Strings[i]) then break ;
    end ;
  end ;
end ;

//***********************************
//*                                 *
//*      GESTION DES MESSAGES       *
//*                                 *
//***********************************
{---------------------------------------------------------------------------------------}
procedure TOF_ConvRubBud.InitMsg ;
{---------------------------------------------------------------------------------------}
var
  Titre: string ;
begin
  LMsg:=THMsgBox.create(FMenuG) ;
  Titre:='Conversion des rubriques budgétaires en rubriques;' ;
  {00} LMsg.Mess.Add('0;'+Titre+'Vous devez renseigner la famille de rubriques;W;O;O;O');
  {01} LMsg.Mess.Add('1;'+Titre+'La rubrique %% existe déjà voulez-vous la remplacer;Q;YNL;Y;Y') ;
  {02} LMsg.Mess.Add('2;'+Titre+'La conversion des rubriques budgétaires a été effectuée;A;O;O;O');
  {03} LMsg.Mess.Add('2;'+Titre+'Vous devez sélectionner au moins un compte de correspondance;W;O;O;O');
end;

{---------------------------------------------------------------------------------------}
function TOF_ConvRubBud.AfficheMsg(num : integer;Av,Ap : string ) : Word ;
{---------------------------------------------------------------------------------------}
begin
  Result:=mrNone ; if
  Num > 3 then Exit ;
  Result:=LMsg.Execute(num,Av,Ap) ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_ConvRubBud.FermeMsg ;
{---------------------------------------------------------------------------------------}
begin
  LMsg.Free ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_ConvRubBud.BHelpClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  CallHelpTopic(Ecran) ;
end;

initialization
  registerclasses([TOF_ConvRubBud]);
end.
