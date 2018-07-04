{ Unité : Création de rubriques à partir des tables libres
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
4.3.1.101.001    17/06/03    JP     Création de l'unité

--------------------------------------------------------------------------------------}
Unit CPRUBTLIB_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Windows,
     {$IFDEF EAGLCLIENT}
     MaineAGL,
     {$ELSE}
     FE_Main, db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     {$ENDIF}
     CPAFFLISTRUB_TOF,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     ent1,
     HMsgBox,
     FAMRUB_TOF,
     UTob,
     UTOF ;

procedure CPLanceFiche_CpRubTLib(vStRange, vStLequel, vStArgs: string);

type
  TOF_CPRUBTLIB = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    { Declarations privées}
    FListe              : THGrid ;
    AuMoinsUneMiseAJour : Boolean ;
    Tous                : Boolean ;
    FirstLoad           : Boolean ;
    LMsg                : THMsgBox ;
    TypeCompte          : THValComboBox ;
    LaFamille           : THedit ;
    Pages               : TPageControl ;
    ListeRub            : Tob ;
    CTab                : array [1..10] of THValComboBox ;
    Tb                  : TOB;

    procedure TableLibreToCombos ;
    procedure TypeCompteToCombo ;
    procedure NettoyerGrille;

    procedure OnBOuvrirClick         (Sender : TObject) ;
    procedure OnBChercherTablesLibres(Sender : TObject) ;
    procedure OnElipsisClickLafamille(Sender : TObject) ;
    procedure OnTypeCompteChange     (Sender : TObject) ;

    procedure InsertOrUpdateRubriques(Nom,Libelle,Famille,TypeRub,Compte1,Exclus1,Axe,TablesLibres : string ) ;
    procedure TraiteTablesLibres     (TablesLibres: array of string ;Libelle : string ) ;
    procedure InitMsg ;
    procedure FermeMsg ;

    function  AfficheMsg (num : Integer; Av, Ap : string ) : Word ;
    function  AutoriseMaj(Nom : string) : Boolean ;
 end ;

Implementation

uses
    {$IFDEF MODENT1}
    CPProcMetier,
    {$ENDIF MODENT1}
    {$IFDEF eAGLCLIENT}
    MenuOLX
    {$ELSE}
    MenuOLG
    {$ENDIF eAGLCLIENT}
    ;

procedure CPLanceFiche_CpRubTLib(vStRange, vStLequel, vStArgs: string);
begin
  AGLLanceFiche('CP', 'CPRUBTLIB', vStRange, vStLequel, vStArgs);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPRUBTLIB.OnLoad ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPRUBTLIB.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
var
  BButton : TButton ;
  i  : integer ;
  HL : TLabel ;
begin
  inherited ;
  //SG6 21/01/05 FQ 15291
  Ecran.HelpContext := 7815000;

  FirstLoad:=True ;

  InitMsg ;

  TypeCompte := THValComboBox(GetControl('TYPECOMPTE')) ;

  if (TypeCompte<>nil) and not assigned(TypeCompte.OnChange) then
    TypeCompte.OnChange := OnTypeCompteChange ;

  TypeCompteToCombo ;

  OnTypeCompteChange(TypeCompte) ;

  HL := tLabel(GetControl('FE__HLABEL2')) ; If HL<>NIL Then HL.Visible:=FALSE ;
  HL := tLabel(GetControl('FE__HLABEL3')) ; If HL<>NIL Then HL.Visible:=FALSE ;

  BButton := TButton(GetControl('BOUVRIR')) ;
  if (BButton <> nil) then BButton.OnClick:=OnBOuvrirClick ;

  BButton := TButton(GetControl('BCHERCHE')) ;
  if (BButton <> nil) then BButton.OnClick := OnBChercherTablesLibres;

  LaFamille := THEdit(GetControl('LAFAMILLE')) ;
  if (LaFamille<>nil) then
    LaFamille.OnElipsisClick := OnElipsisClickLafamille ;

  Pages := TPageControl(GetControl('PAGES')) ;

  FListe := THGrid(GetControl('FListe1',True)) ;

  for i:=1 to 10 do
    CTab[i] := THValComboBox(GetControl('TAB'+IntToStr(i))) ;
  TableLibreToCombos ;

  ListeRub := Tob.Create('§LISTERUB',nil,-1) ;

  OnBChercherTablesLibres(Self);
  FirstLoad:=False;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPRUBTLIB.OnClose ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
  FermeMsg ;
  FreeAndNil(ListeRub);
  FreeAndNil(Tb);
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPRUBTLIB.TraiteTablesLibres(TablesLibres: array of string ;Libelle : string ) ;
{---------------------------------------------------------------------------------------}
var
  i        : integer ;
  Virgule,
  Nom,
  Famille,
  TypeRub,
  Compte1,
  Exclus1,
  Axe      : string ;
begin
  Axe     := '';
  Compte1 := '';
  Exclus1 := '';
  Famille := LaFamille.text ;
  TypeRub := 'GEN' ;

  if (TypeCompte.ItemIndex<>0) and (TypeCompte.ItemIndex<>1) then begin
    TypeRub := 'ANA';
    Axe     := 'A' + IntToStr(TypeCompte.ItemIndex - 1);
  end ;

  Nom := Axe + ':';
  for i := 0 to 9 do Nom := Nom + TablesLibres[i]; //de 0..9 comme les tables libres!
  Virgule := '' ;

  for i:=0 to 9 do begin
    if (TablesLibres[i]='') then Compte1 := Compte1 + Virgule + '#'
                            else Compte1 := Compte1 + Virgule + TablesLibres[i] ;
    Virgule:=',' ;
  end ;

  if Compte1 <> '' then
    Compte1 := Compte1 + '(SM)' ;

  {Insertion des rubriques dans la table}
  InsertOrUpdateRubriques(Nom,Libelle,Famille,TypeRub,Compte1,Exclus1,Axe,'X') ;
end ;

{Ecriture des rubriques crées dans la table "Rubrique"
{---------------------------------------------------------------------------------------}
procedure TOF_CPRUBTLIB.InsertOrUpdateRubriques(Nom,Libelle,Famille,TypeRub,Compte1,Exclus1,Axe,TablesLibres : string ) ;
{---------------------------------------------------------------------------------------}
var
  Q      : TQuery ;
  Okok,
  InsOk  : Boolean ;
  T1,
  TRub   : TOB ;
const
  Prefix = 'TL:';
begin
  if Nom <> '' then begin
    Nom := Prefix + copy(Nom, 1, 14);

    Q := OpenSql('SELECT COUNT(RB_RUBRIQUE) FROM RUBRIQUE WHERE RB_RUBRIQUE="'+Nom+'"',True) ;
    InsOk :=  Q.Fields[0].AsInteger = 0;
    Ferme(Q);
    TRub := TOB.Create('RUBRIQUE', nil, -1);

    if InsOk then Okok := True
             else Okok := AutoriseMaj(Nom);

    if Okok then begin
      AuMoinsUneMiseAJour:=True ;
      TRub.PutValue('RB_RUBRIQUE'  , Nom );
      TRub.PutValue('RB_LIBELLE'   , Libelle);
      TRub.PutValue('RB_FAMILLES'  , Famille);
      TRub.PutValue('RB_SIGNERUB'  , 'POS'  );
      TRub.PutValue('RB_TYPERUB'   , TypeRub);
      TRub.PutValue('RB_COMPTE1'   , Compte1);
      TRub.PutValue('RB_EXCLUSION1', Exclus1);
      TRub.PutValue('RB_AXE'       , Axe    );
      TRub.PutValue('RB_NATRUB'    , 'CPT'  );
      TRub.PutValue('RB_CLASSERUB' , 'TLI'  );
      TRub.PutValue('RB_CODEABREGE', copy(Nom,4,6));
      TRub.PutValue('RB_TABLELIBRE', TablesLibres) ;
      TRub.PutValue('RB_PREDEFINI' , 'DOS'   );
      TRub.PutValue('RB_NODOSSIER' , '000000');
      TRub.PutValue('RB_NODOSSIER', V_PGI.NoDossier);
      if InsOk then TRub.InsertDb(nil)
               else TRub.UpdateDB;

      T1 := Tob.Create('$Fille',ListeRub,-1) ;
      T1.AddChampSup('RS_CODE',True) ;
      T1.PutValue('RS_CODE',Nom) ;
      T1.AddChampSup('RS_LIBELLE',True) ;
      T1.PutValue('RS_LIBELLE',Libelle) ;
      T1.AddChampSup('RS_COMMENT',True) ;
      if InsOk then T1.PutValue('RS_COMMENT', 'Rubrique créée')
               else T1.PutValue('RS_COMMENT', 'Rubrique mise à jour');
    end ;
    FreeAndNil(TRub) ;
  end ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPRUBTLIB.InitMsg ;
{---------------------------------------------------------------------------------------}
begin
  LMsg:=THMsgBox.create(FMenuG);
  {00} LMsg.Mess.Add('0;Conversion rubriques;Le champ QUELCONVERSION n''existe pas dans le fenetre;W;O;O;O') ;
  {01} LMsg.Mess.Add('1;Conversion des tables libres en rubriques;La rubrique %% existe déjà voulez-vous la remplacer;Q;YNL;Y;Y') ;
  {02} LMsg.Mess.Add('2;Conversion des tables libres en rubriques;Vous devez renseigner la famille de rubriques;W;O;O;O');
  {03} LMsg.Mess.Add('2;Conversion des tables libres en rubriques;Vous devez sélectionner au moins une valeur de tables libres;W;O;O;O');
  {04} LMsg.Mess.Add('2;Conversion des tables libres en rubriques;La création des rubriques a été effectuée;A;O;O;O');
  {05} LMsg.Mess.Add('2;Conversion des tables libres en rubriques;;A;O;O;O');
end;

{---------------------------------------------------------------------------------------}
function TOF_CPRUBTLIB.AfficheMsg(num : integer;Av,Ap : string ) : Word ;
{---------------------------------------------------------------------------------------}
begin
  Result := mrNone ;
  if Num > 5 then Exit ;
  Result := LMsg.Execute(num,Av,Ap) ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPRUBTLIB.FermeMsg ;
{---------------------------------------------------------------------------------------}
begin
  LMsg.Free ;
end;

{Si la rubrique existe déjà, doit-on la mettre à jour ?
{---------------------------------------------------------------------------------------}
function TOF_CPRUBTLIB.AutoriseMaj(Nom : String) : boolean ;
{---------------------------------------------------------------------------------------}
var
  Ret : Word ;
begin
  if not Tous then begin
    Ret:=AfficheMsg(1, Nom, '') ;
    if Ret = mrAll then Tous := True ;
    Result:= (Ret = mrYes) or (Ret = mrAll);
  end
  else
    Result:=True ;
end ;

{Recherche des tables libres en fonction du type selectionné
{---------------------------------------------------------------------------------------}
procedure TOF_CPRUBTLIB.OnBChercherTablesLibres(Sender : TObject) ;
{---------------------------------------------------------------------------------------}
var
  QQ                       : TQuery;
  i,j, NombreChamps          : Integer ;
  Axe, Prefix, Table,
  Nature, PointVirgule,
  Virgule, LePipe, LeAnd,
  Sql, LeMultiF, LesChamps,
  ConcatNom, ConcatNat, SommeNom,
  LeftJoin, Where, GroupBy : String ;
begin
  if FirstLoad then Exit ;

  case TypeCompte.ItemIndex of
    0 :  begin Prefix:='G_' ; Nature:='G0' ; Table:='GENERAUX' ; end ;
    1 :  begin Prefix:='T_' ; Nature:='T0' ; Table:='TIERS' ;    end ;
    else begin Prefix:='S_' ; Nature:='S0';  Table:='SECTION' ;  end ;
  end ;

  SommeNom:='' ;
  ConcatNom:='' ;
  Virgule:=' ' ;
  LePipe:=' ' ;
  NombreChamps:=0 ;
  Axe:='' ;

  if (TypeCompte.ItemIndex<>0) and (TypeCompte.ItemIndex<>1) then
    Axe := 'A' + IntToStr(TypeCompte.ItemIndex-1) ;

  { YMO 01/02/2006 FQ17389 Dorénavant, les niveaux sont dans l'ordre, et le code est le fruit de la concaténation
   des codes des tables libres, mais dans l'ordre ! on ne peut plus créer une rubrique CODE0CODE1 ert
   une rubrique CODE1CODE0}
  for j:=1 to 10 do
  begin
    for i:=1 to 10 do

    if (CTab[i]<>nil) and (CTab[i].ItemIndex=j) then begin
      LeMultiF := LeMultiF  + PointVirgule + Prefix + 'TABLE' + CTab[i].Value;
      LesChamps:= LesChamps + Virgule      + Prefix + 'TABLE' + CTab[i].Value;
      ConcatNom:= ConcatNom + LePipe       + Prefix + 'TABLE' + CTab[i].Value;
      SommeNom := SommeNom + ', '       + Prefix + 'TABLE' + CTab[i].Value;
      ConcatNat:= ConcatNat + LePipe + 'NT' + IntToStr(i) + '.NT_LIBELLE' ;

      LeftJoin := LeftJoin  + ' LEFT JOIN NATCPTE NT' + IntToStr(i) + ' ON NT' + IntToStr(i) +
                              '.NT_NATURE=' + Prefix + 'TABLE' + CTab[i].Value ;

      Where    := Where     + LeAnd + Prefix + 'TABLE' + CTab[i].Value + '<>""' + ' AND NT' +
                              IntToStr(i) + '.NT_TYPECPTE="' + Nature + CTab[i].Value + '" ';

      GroupBy  := GroupBy   + Virgule + Prefix + 'TABLE' + CTab[i].Value + ',' + 'NT' +
                              IntToStr(i) + '.NT_LIBELLE' ;
      inc(NombreChamps) ;
      PointVirgule:=';' ; Virgule:=',' ; LePipe:='||' ;LeAnd:=' AND ' ;

    Break; // pas deux fois le même
    end ;
  end ;

  if NombreChamps>0 then begin
    {JP 18/06/03 : le code suivant ne donne pas le résultat voulu
    Sql := 'SELECT '+LesChamps;
    if ConcatNom<>'' then Sql:=Sql+','+ConcatNom+' AS Code' ;
    if ConcatNat<>'' then Sql:=Sql+','+ConcatNat+' AS Libelle' ; }

    Sql := 'SELECT ' + ConcatNom + ' AS Code' ;
    Sql := Sql + ',' + ConcatNat + ' AS Libelle' ;
    Sql := Sql + SommeNom  ;
    Sql := Sql + ' FROM ' + Table+LeftJoin + ' WHERE ' + Where ;
    if Axe <> '' then Sql := Sql + LeAnd + ' ' + Prefix + 'AXE="' + Axe + '" ' ;
    Sql := Sql + ' GROUP BY '+ GroupBy ;

    NettoyerGrille;
    if Tb = nil then
      Tb := TOB.Create('', nil, -1);
    QQ := OpenSQL(Sql, True);
    if not QQ.EOF then
    begin
      Tb.LoadDetailDB('', '', '', QQ, False);
      Tb.PutGridDetail(Fliste, False, False, '');
    end;
    Ferme(QQ);
  end;
  
end ;

{Lancement de la création ou de la mise à jour des rubriques
{---------------------------------------------------------------------------------------}
procedure TOF_CPRUBTLIB.OnBouvrirClick(Sender : TObject) ;
{---------------------------------------------------------------------------------------}
var
  TablesLibres : array[0..9] of string ;
  i, n        : Integer;
  Prefix, NumTL,
  Libelle      : string ;

begin
  AuMoinsUneMiseAJour:=False ;
  if FListe.nbSelected < 1 then begin
    AfficheMsg(3,'','') ;
    Exit ;
  end ;

  Tous:=False ;
  if FListe.nbSelected > 0 then
    for n := 0 to FListe.nbSelected - 1 do begin
      if LaFamille.Text = '' then begin
        AfficheMsg(2,'','') ;
        Pages.ActivePage := Pages.Pages[1] ;
        LaFamille.SetFocus ; Exit ;
      end ;

      FListe.GotoLeBookmark(n) ;

      case TypeCompte.ItemIndex of
        0 :  Prefix:='G_' ;
        1 :  Prefix:='T_' ;
        else Prefix:='S_' ;
      end ;

      { YMO 01/02/2006 FQ17389 Dorénavant, les niveaux sont dans l'ordre, et le code est le fruit de la concaténation
      des codes des tables libres, mais dans l'ordre ! on ne peut plus créer une rubrique CODE0CODE1 ert
      une rubrique CODE1CODE0}
      for i:=1 to 10 do
      begin
        if CTab[i].ItemIndex = 0 then continue;
        NumTL:=CTab[i].Values[CTab[i].ItemIndex];
        TablesLibres[StrToInt(NumTL)] := '';
        if (CTab[i]<>nil) and (NumTL<>'') then
        //  TablesLibres[StrToInt(CTab[i].Values[CTab[i].ItemIndex])] := Tb.Detail[FListe.Row - 1].GetValue('CODE');

        TablesLibres[StrToInt(NumTL)] := Tb.Detail[FListe.Row - 1].GetValue(Prefix+'TABLE'+NumTL);
      end;

      Libelle := Tb.Detail[FListe.Row - 1].GetValue('LIBELLE');
      TraiteTablesLibres(TablesLibres,Libelle) ;
    end ;

  FListe.ClearSelected;
  if AuMoinsUneMiseAJour then begin
    AfficheListe(ListeRub,'Liste des rubriques integrées',['Nom Rubrique','Libellé']) ;
    AfficheMsg(4,'','') ;
    ListeRub.ClearDetail;
  end
end ;

{Appel de l'écran de sélection des familles de rubriques : FAMRUB_TOF
{---------------------------------------------------------------------------------------}
procedure TOF_CPRUBTLIB.OnElipsisClickLafamille(Sender : TObject) ;
{---------------------------------------------------------------------------------------}
var
  Fam,LibFam : string ;
begin
  Fam := LaFAmille.Text ;
  ParametrageFamilleRubrique('','',Fam,LibFam,False,TRUE) ;
  {if Fam <> '' then} LaFAmille.Text := Fam ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPRUBTLIB.OnTypeCompteChange(Sender : TObject) ;
{---------------------------------------------------------------------------------------}
begin
  // YMO 01/02/2006 Maj
  if (FListe <> nil) then
    NettoyerGrille;
  if (TypeCompte <> nil) then
    TableLibreToCombos ;
end ;

{Remplissage des combos contenant les tables libres en fonction du type de compte sélectionné
{---------------------------------------------------------------------------------------}
procedure TOF_CPRUBTLIB.TableLibreToCombos ;
{---------------------------------------------------------------------------------------}
var
  ValListe: HTStringList ;
  i,j,nbOK: integer ;
  StTmp,NomTable,OkTable: string ;
begin
  NbOK:=0 ;
  for i:=1 to 10 do
    if CTab[i] <> nil then begin
      CTab[i].Items.Clear ;
      CTab[i].Values.Clear ;
      CTab[i].Items.Add('<<Aucune>>') ;
      CTab[i].Values.Add('') ;
      CTab[i].ItemIndex:=0 ;
    end ;

  ValListe:=HTStringList.Create ;

  if typeCompte.Text<>'' then
    GetLibelleTableLibre(copy(typeCompte.Value,1,1),ValListe) ;

  for j:=0 to ValListe.Count-1 do begin
    StTmp:=ValListe.Strings[j] ;
    NomTable:=ReadTokenSt(StTmp) ;
    OkTable:=ReadTokenSt(StTmp) ;
    if OkTable = 'X' then begin
      for i:=1 to 10 do
        if CTab[i]<>nil then begin
          CTab[i].Items.Add(NomTable) ;
          CTab[i].Values.Add(IntToStr(j)) ;
        end ;
      inc(NbOK) ;
    end ;
  end ;

  for i:=1 to 10 do
    if (CTab[i]<>nil) then begin
      if (NbOk<i) then
        CTab[i].Enabled:=False
      else
        CTab[i].Enabled:=True ;
    end;

  {$IFDEF CCS3}
  For i:=5 To 10 Do if CTab[i]<>nil then CTab[i].Visible:=FALSE ;
  {$ENDIF}
  ValListe.Free ;
end ;

{Remplissage du combo TypeCompte en ouverture de la fiche
{---------------------------------------------------------------------------------------}
procedure TOF_CPRUBTLIB.TypeCompteToCombo ;
{---------------------------------------------------------------------------------------}
var
  TabValue : array [1..3,1..7] of string ;
  i : integer ;
begin
  if (TypeCompte=nil) or not (TypeCompte is THValComboBox) then Exit ;

  TabValue[1,1]:='RUG'; TabValue[1,2]:='RUT';
  for i:=1 to 5 do TabValue[1,i+2]:='RU'+IntToStr(i) ;

  TabValue[2,1]:='GE';  TabValue[2,2]:='AU';
  for i:=1 to 5 do TabValue[2,i+2]:='A'+IntToStr(i) ;

  TabValue[3,1]:='GE'; TabValue[3,2]:='TI';
  for i:=1 to 5 do TabValue[3,i+2]:='S'+IntToStr(i) ;

  TypeCompte.Clear ;
  TypeCompte.Items.Add('Généraux') ;
  TypeCompte.Values.Add(TabValue[3, 1]) ;
  TypeCompte.Items.Add('Auxiliaire') ;
  TypeCompte.Values.Add(TabValue[3, 2]) ;
  TypeCompte.Items.Add('Section axe 1') ;
  TypeCompte.Values.Add(TabValue[3, 3]) ;

  {$IFNDEF CCS3}
  TypeCompte.Items.Add('Section axe 2') ;
  TypeCompte.Values.Add(TabValue[3, 4]) ;
  {$ENDIF}

  // JLD5Axes
  if not EstSerie(S3) then begin
    TypeCompte.Items.Add('Section axe 3') ;
    TypeCompte.Values.Add(TabValue[3, 5]) ;
    TypeCompte.Items.Add('Section axe 4') ;
    TypeCompte.Values.Add(TabValue[3, 6]) ;
    TypeCompte.Items.Add('Section axe 5') ;
    TypeCompte.Values.Add(TabValue[3, 7]) ;
  end ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPRUBTLIB.NettoyerGrille;
{---------------------------------------------------------------------------------------}
var
  n, NbLignes : Integer;

begin
NbLignes:=FListe.RowCount;
// YMO 01/02/2006 FQ17389
FListe.Rows[1].Clear;
for n := 1 to NbLignes do FListe.DeleteRow(2);

end;

Initialization
  registerclasses ( [ TOF_CPRUBTLIB ] ) ;
end.

