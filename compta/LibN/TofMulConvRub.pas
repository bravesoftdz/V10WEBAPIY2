unit TofMulConvRub;

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
      {$IFDEF MODENT1}
      CPTypeCons,
      {$ENDIF MODENT1}
      Forms,Controls,StdCtrls,Graphics,Classes,sysutils,comctrls,
      HCtrls,ent1,HEnt1,HMsgBox,UTOF,UTOB,LookUp,HDB,Vierge,
      HRichOLE,UtilDiv,Hqry,choix,
      HTB97 ;

type TOF_MulCorRub = class (TOF)
     private
       LaFamille: THedit ;
       ListeRub: Tob ;
       TypeCompte,PlanCompte : THValComboBox ;
       //CTab: array [1..10] of THValComboBox ;
       RuNatureRupt,RuPlanRupt,CrType: TEdit ;
       Pages: TPageControl ;
      {$IFDEF EAGLCLIENT}
          FListe              : THGrid ;
      {$ELSE}
          FListe              : THDBGrid ;
      {$ENDIF}
       QuelConversion: integer ; // 1 : rupture, 2 : compte correspondance
       ModeSolde: string ;
       Tous,FirstLoad,AuMoinsUneMiseAJour: boolean ;
       LMsg : THMsgBox ;
       BHelp: TToolbarButton97;
//       procedure TableLibreToCombos ;
       procedure PlanCompteToCombo ;
       procedure TypeCompteToCombo ;
       procedure OnBZoomClick(Sender : TObject) ;
       procedure OnTypeCompteChange(Sender : TObject) ;
       procedure OnPlanCompteChange(Sender : TObject) ;
       procedure OnBOuvrirClick(Sender : TObject) ;
       procedure OnElipsisClickLafamille(Sender : TObject) ;
       procedure TraiteRubrique(Nature,Plan,Classe,Libelle : string) ;
       function  RecupAxe(Nature : string ;var Axe : string;var LonAxe : integer;var NumAxe :  TFichierBase ) : boolean;
       procedure TraiteCorresp(TypeCor,Corresp,Libelle : string ) ;
       function  GenereListePlusCompacte(Corresp: string ; Liste: HTStrings;var Compte,Exclus: string): boolean ;
       procedure InsertOrUpdateRubriques(Nom,Libelle,Famille,TypeRub,Compte1,Exclus1,Axe,TablesLibres : string ) ;
       function AutoriseMaj(Nom : String) : boolean ;
       procedure InitMsg ;
       function AfficheMsg(num : integer;Av,Ap : string ) : Word ;
       procedure FermeMsg ;
       procedure BHelpClick(Sender: TObject);
     public
       procedure OnLoad ; override ;
       procedure OnClose ; override ;
       procedure OnArgument (stArgument : String )  ; override ;
     end;

implementation

uses
    {$IFDEF eAGLCLIENT}
    MenuOLX, FamRub_TOF
    {$ELSE}
    MenuOLG,
     //{$IFDEF ESP}
     FamRub_TOF
     (*{$ELSE}
     FamRub
     {$ENDIF ESP}*)
    {$ENDIF eAGLCLIENT}
    ;
(*
{---------------------------------------------------------------------------------------}
procedure TOF_MulCorRub.TableLibreToCombos ;
{---------------------------------------------------------------------------------------}
var
  ValListe: HTStringList ;
  i,j,nbOK: integer ;
  StTmp,NomTable,OkTable: string ;
begin
  NbOK:=0 ;
  for i:=1 to 10 do
    if CTab[i] <> nil then
    begin
      CTab[i].Items.Clear ;
      CTab[i].Values.Clear ;
      CTab[i].Items.Add('<<Aucune>>') ;
      CTab[i].Values.Add('') ;
      CTab[i].ItemIndex:=0 ;
    end ;

  ValListe:=HTStringList.Create ;

  if typeCompte.Text<>'' then
    GetLibelleTableLibre(copy(typeCompte.Value,1,1),ValListe) ;

  for j:=0 to ValListe.Count-1 do
  begin
    StTmp:=ValListe.Strings[j] ;
    NomTable:=ReadTokenSt(StTmp) ;
    OkTable:=ReadTokenSt(StTmp) ;
    if OkTable='X' then
    begin
      for i:=1 to 10 do
        if CTab[i]<>nil then
        begin
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
  *)
{---------------------------------------------------------------------------------------}
procedure TOF_MulCorRub.PlanCompteToCombo ;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery ;
begin
  if (PlanCompte=nil) or not (PlanCompte is THValComboBox) then Exit ;

  {JP : PlanCompte.Clear ne nettoie pas Values}
  PlanCompte.Items.Clear;
  PlanCompte.Values.Clear;

  case QuelConversion of
    1 : begin
          Q:=OpenSql('SELECT CC_CODE,CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="'+TypeCompte.Value+'"',True) ;
          While not Q.Eof do begin
            PlanCompte.Items.Add(Q.Findfield('CC_LIBELLE').AsString) ;
            PlanCompte.Values.Add(Q.Findfield('CC_CODE').AsString) ;
            Q.next ;
          end ;
          ferme(Q) ;
        end ;
    2 : begin
          PlanCompte.Items.Add('Plan1') ;
          PlanCompte.Values.Add('1') ;
          if not((EstSerie(S3)) or (EstSerie(S5))) then begin
            PlanCompte.Items.Add('Plan2') ;
            PlanCompte.Values.Add('2');
          end ;
        end;
  end ;
end ;

// Quoi : 1 :
{---------------------------------------------------------------------------------------}
procedure TOF_MulCorRub.TypeCompteToCombo ;
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
  TypeCompte.Values.Add(TabValue[QuelConversion,1]) ;
  TypeCompte.Items.Add('Auxiliaire') ;
  TypeCompte.Values.Add(TabValue[QuelConversion,2]) ;
  TypeCompte.Items.Add('Section axe 1') ;
  TypeCompte.Values.Add(TabValue[QuelConversion,3]) ;

  {$IFNDEF CCS3}
  TypeCompte.Items.Add('Section axe 2') ;
  TypeCompte.Values.Add(TabValue[QuelConversion,4]) ;
  {$ENDIF}

  // JLD5Axes
  if not EstSerie(S3) then begin
    TypeCompte.Items.Add('Section axe 3') ;
    TypeCompte.Values.Add(TabValue[QuelConversion,5]) ;
    TypeCompte.Items.Add('Section axe 4') ;
    TypeCompte.Values.Add(TabValue[QuelConversion,6]) ;
    TypeCompte.Items.Add('Section axe 5') ;
    TypeCompte.Values.Add(TabValue[QuelConversion,7]) ;
  end ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_MulCorRub.OnArgument (stArgument : String )  ;
{---------------------------------------------------------------------------------------}
var
  BButton : TButton ;
  Quel    : TEdit ;
  HL : TLabel ;
begin
  inherited ;
  ModeSolde:='(SM)' ;
  FirstLoad:=True ;
  Quel := TEdit(GetControl('QUELCONVERSION')) ;

  if Quel = nil then begin
    AfficheMsg(0,'','') ;
    Exit;
  end ;

  QuelConversion := StrToInt(Quel.Text) ;

  // Aide
  case QuelConversion of
     1 : TFMul(Ecran).HelpContext:=7813000;
     2 : TFMul(Ecran).HelpContext:=7814000;
     3 : TFMul(Ecran).HelpContext:=7815000;
  end ;

  if (QuelConversion < 1) or (QuelConversion > 2) then
    QuelConversion:=1 ;

  InitMsg ;
  {gestion des diverses infos du multicritères communes}
  TypeCompte:=THValComboBox(GetControl('TYPECOMPTE')) ;

  if (TypeCompte<>nil) and not assigned(TypeCompte.OnChange) then
    TypeCompte.OnChange := OnTypeCompteChange ;

  PlanCompte:=THValComboBox(GetControl('PLANCOMPTE')) ;
  if (PlanCompte<>nil) and not assigned(PlanCompte.OnChange) then
    PlanCompte.OnChange:=OnPlanCompteChange ;

  TypeCompteToCombo ;
  OnTypeCompteChange(TypeCompte) ;
  PlanCompteToCombo ;

  HL:=tLabel(GetControl('FE__HLABEL2')) ; If HL<>NIL Then HL.Visible:=FALSE ;
  HL:=tLabel(GetControl('FE__HLABEL3')) ; If HL<>NIL Then HL.Visible:=FALSE ;

  BButton:=TButton(GetControl('BOUVRIR')) ;
  if (BButton<>nil) then BButton.OnClick:=OnBOuvrirClick ;

  BButton:=TButton(GetControl('BZOOM')) ;
  if (BButton<>nil) and not assigned(BButton.OnClick) then BButton.OnClick:=OnBZoomClick ;

  LaFamille:=THEdit(GetControl('LAFAMILLE')) ;
  if (LaFamille<>nil) and not assigned(LaFamille.OnElipsisClick) then
    LaFamille.OnElipsisClick:=OnElipsisClickLafamille ;

  Pages:=TPageControl(GetControl('PAGES')) ;

  {$IFDEF EAGLCLIENT}
    FListe := THGrid(GetControl('FListe',True)) ;
  {$ELSE}
    FListe := THDBGrid(GetControl('FListe',True)) ;
  {$ENDIF}

  {gestion des diverses infos du multicritères spécifiques}
  case QuelConversion of
    1 : begin
          RuNatureRupt:=TEdit(GetControl('RU_NATURERUPT')) ;
          if RuNatureRupt <> nil then
            RuNatureRupt.Text:=W_W ;
          RuPlanRupt:=TEdit(GetControl('RU_PLANRUPT')) ;
          if RuPlanRupt <> nil then
            RuPlanRupt.Text:=W_W ;
        end ;
    2 : begin
          CrType:=TEdit(GetControl('CR_TYPE')) ;
          if CrType <> nil then
            CrType.Text:=W_W ;
        end ;
   end ;
  ListeRub:=Tob.Create('§LISTERUB',nil,-1) ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_MulCorRub.OnLoad;
{---------------------------------------------------------------------------------------}
begin
  inherited ;

  BHelp:=TToolbarButton97(GetControl('BAIDE')) ;
  if (BHelp<>nil) and (not Assigned(BHelp.OnClick)) then BHelp.OnClick:=BHelpClick ;
  FirstLoad:=False ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_MulCorRub.OnClose;
{---------------------------------------------------------------------------------------}
begin
  inherited ;
  FermeMsg ;
  ListeRub.Free ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_MulCorRub.OnTypeCompteChange(Sender : TObject) ;
{---------------------------------------------------------------------------------------}
begin
  if (TypeCompte=nil) then Exit ;
  case QuelConversion of
    1 : begin
          PlanCompteToCombo ;
          RuNatureRupt.Text:=TypeCompte.Value ;
        end ;
    2 : CrType.Text:=TypeCompte.Value+PlanCompte.Value ;
  end ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_MulCorRub.OnPlanCompteChange(Sender : TObject) ;
{---------------------------------------------------------------------------------------}
begin
  if (PlanCompte=nil) then Exit ;
  case QuelConversion of
    1 : RuPlanRupt.Text := PlanCompte.Value;
    2 : CrType.Text:=TypeCompte.Value+PlanCompte.Value ;
  end ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_MulCorRub.OnBZoomClick(Sender : TObject) ;
{---------------------------------------------------------------------------------------}
begin
  {Appel de la tot CPRUBFAMILLE : UTOTRUBFAMILLE}
  ParamTable('CPRUBFAMILLE', taCreat, 7775000,nil, 3,'Familles de rubriques') ;
  PlanCompteToCombo ;
end ;

{Appel de l'écran de sélection des familles de rubriques : FAMRUB_TOF
{---------------------------------------------------------------------------------------}
procedure TOF_MulCorRub.OnElipsisClickLafamille(Sender : TObject) ;
{---------------------------------------------------------------------------------------}
Var
  Fam,LibFam : String ;
begin
  Fam := LaFAmille.Text ;
  ParametrageFamilleRubrique('','',Fam,LibFam,False,TRUE) ;
  {if Fam <> '' Then} LaFAmille.Text := Fam ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_MulCorRub.OnBouvrirClick(Sender : TObject) ;
{---------------------------------------------------------------------------------------}
var
  Q : THQuery ;
  n : integer;
  Plan,
  Classe,
  TypeCor,
  Corresp,
  Nature,
  Libelle  : string ;
begin
  AuMoinsUneMiseAJour:=False ;
  if FListe.nbSelected < 1 then begin
    AfficheMsg(3,'','') ;
    Exit ;
  end;

  {$IFDEF EAGLCLIENT}
    Q := TFMul(Ecran).Q;
  {$ELSE}
    Q := THQuery(FListe.DataSource.DataSet) ;
  {$ENDIF}

  Tous:=False ;
  for n := 0 to  FListe.nbSelected - 1 do begin
    FListe.GotoLeBookmark(n) ;
    {$IFDEF EAGLCLIENT}
    Q.TQ.Seek(FListe.Row - 1);
    {$ENDIF}

    if LaFamille.Text='' then begin
      AfficheMsg(2,'','') ;
      Pages.ActivePage:=Pages.Pages[3] ; LaFamille.SetFocus ; Exit ;
    end ;

    case QuelConversion of
      1 : begin
            Nature :=Q.Findfield('RU_NATURERUPT').AsString ;
            Plan   :=Q.Findfield('RU_PLANRUPT').AsString ;
            Classe :=Q.Findfield('RU_CLASSE').AsString ;
            Libelle:=Q.Findfield('RU_LIBELLECLASSE').AsString ;
            TraiteRubrique(Nature,Plan,Classe,Libelle) ;
          end ;
      2 : begin
            TypeCor:=Q.Findfield('CR_TYPE').AsString ;
            Corresp:=Q.Findfield('CR_CORRESP').AsString ;
            Libelle:=Q.Findfield('CR_LIBELLE').AsString ;
            TraiteCorresp(TypeCor,Corresp,Libelle) ;
          end ;
    end ;
  end ;
  FListe.ClearSelected;

  if AuMoinsUneMiseAJour then begin
    AfficheListe(ListeRub,'Liste des rubriques integrées',['Nom Rubrique','Libellé']) ;
    AfficheMsg(4,'','') ;
    ListeRub.ClearDetail;
  end
end ;

{Ecriture des rubriques crées dans la table "Rubrique"
{---------------------------------------------------------------------------------------}
procedure TOF_MulCorRub.InsertOrUpdateRubriques(Nom,Libelle,Famille,TypeRub,Compte1,Exclus1,Axe,TablesLibres : string ) ;
{---------------------------------------------------------------------------------------}
var
  Q      : TQuery ;
  Okok,
  InsOk  : Boolean ;
  T1,
  TRub   : TOB ;
  Prefix : string;
begin
  case QuelConversion of
    1 : Prefix:='RU:' ;
    2 : Prefix:='CO:' ;
  end ;

  Nom := Prefix + Copy(Nom, 1, 14);
  if Nom <> '' then begin
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
      TRub.PutValue('RB_CODEABREGE', Copy(Nom,4,6));
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
function TOF_MulCorRub.AutoriseMaj(Nom : String) : boolean ;
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

                             {*      GESTION DES MESSAGES       *}

{---------------------------------------------------------------------------------------}
procedure TOF_MulCorRub.InitMsg ;
{---------------------------------------------------------------------------------------}
begin
  LMsg:=THMsgBox.create(FMenuG) ;
  {00} LMsg.Mess.Add('0;Conversion rubriques;Le champ QUELCONVERSION n''existe pas dans le fenetre;W;O;O;O') ;
  case QuelConversion of
    1 :  begin
          {01}LMsg.Mess.Add('1;Conversion des ruptures en rubriques;La rubrique %% existe déjà voulez-vous la remplacer;Q;YNL;Y;Y') ;
          {02}LMsg.Mess.Add('2;Conversion des ruptures en rubriques;Vous devez renseigner la famille de rubriques;W;O;O;O') ;
          {03}LMsg.Mess.Add('2;Conversion des ruptures en rubriques;Vous devez sélectionner au moins une rupture  ;W;O;O;O') ;
          {04}LMsg.Mess.Add('2;Conversion des ruptures en rubriques;La création des rubriques a été effectuée;A;O;O;O') ;
          {05}LMsg.Mess.Add('2;Conversion des ruptures en rubriques;;A;O;O;O') ;
        end ;
    2 : begin
          {01} LMsg.Mess.Add('1;Conversion des comptes de correspondances en rubriques;La rubrique %% existe déjà voulez-vous la remplacer;Q;YNL;Y;Y') ;
          {02} LMsg.Mess.Add('2;Conversion des comptes de correspondances en rubriques;Vous devez renseigner la famille de rubriques;W;O;O;O');
          {03} LMsg.Mess.Add('2;Conversion des comptes de correspondances en rubriques;Vous devez sélectionner au moins un compte de correspondance;W;O;O;O');
          {04} LMsg.Mess.Add('2;Conversion des comptes de correspondances en rubriques;La création des rubriques a été effectuée;A;O;O;O');
          {05} LMsg.Mess.Add('2;Conversion des comptes de correspondances en rubriques;Le compte de correspondance %% ne possède aucun compte affecté;A;O;O;O');
        end ;
    end ;
  end;

{---------------------------------------------------------------------------------------}
function TOF_MulCorRub.AfficheMsg(num : integer;Av,Ap : string ) : Word ;
{---------------------------------------------------------------------------------------}
begin
  Result := mrNone ;
  if Num > 5 then Exit ;
  Result := LMsg.Execute(num,Av,Ap) ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_MulCorRub.FermeMsg ;
{---------------------------------------------------------------------------------------}
begin
  LMsg.Free ;
end;

{TRAITEMENT DES RUPTURES}
{---------------------------------------------------------------------------------------}
procedure TOF_MulCorRub.TraiteRubrique(Nature,Plan,Classe,Libelle : string) ;
{---------------------------------------------------------------------------------------}
var
  Nom,Famille,
  TypeRub, Compte1,
  Exclus1,Axe,
  PlanSs,Liste : string ;

  LonAxe,i,
  PlanDeb,
  PlanLon      : Integer ;

  NumAxe       : TFichierBase ;
  Q            : TQuery ;
begin
  Nom := Nature + Plan + Classe ;
  Famille := LaFamille.Text;
  Exclus1 := '';

  if (TypeCompte.ItemIndex = 0) or (TypeCompte.ItemIndex = 1) then begin
    TypeRub:='GEN';
    Axe:='' ;
    Compte1:=copy(Classe,1,length(Classe)-1)+ModeSolde ;
  end

  else begin
    if not RecupAxe(Nature,Axe,LonAxe,NumAxe) then Exit ;
    TypeRub:='ANA';
    Q:=OpenSql('SELECT * FROM CHOIXCOD WHERE CC_TYPE="'+Nature+'" AND CC_CODE="'+Plan+'"',True) ;

    if not Q.Eof then Liste:=Q.findfield('CC_LIBRE').AsString
                 else Liste:='' ;
    ferme(Q) ;

    if Liste='' then
      Compte1 := Copy(Classe,1,length(Classe)-1)

    else begin
      memset(Compte1,'?',LonAxe) ;
      PlanSs:=ReadTokenSt(Liste) ;

      while PlanSs<>'' do begin
        //RecupSousSection(NumAxe,PlanSs,PlanDeb,PlanLon) ;
        PlanDeb := 0;
        PlanLon := 0;

        for i:=1 to MaxSousPlan do
          if VH^.SousPlanAxe[NumAxe,i].Code=PlanSs then begin
            PlanDeb := VH^.SousPlanAxe[NumAxe,i].Debut ;
            PlanLon := VH^.SousPlanAxe[NumAxe,i].Longueur ;
            break ;
          end ;
        for i:=1 to PlanLon do Compte1[PlanDeb - 1 + i] := Classe[i];

        PlanSs:=ReadTokenSt(Liste) ;
      end ;
      for i:=LonAxe downto 0 do
        if Compte1[i]='?' then Compte1[i]:=' '
                          else break ;
    end;{if Liste = '' else begin}

    if Compte1 <> '' then
      Compte1 := Trim(Compte1) + ModeSolde ;
  end; {if (TypeCompte. else begin}

  if Compte1 <> '' then
    InsertOrUpdateRubriques(Nom,Libelle,Famille,TypeRub,Compte1,Exclus1,Axe,'-') ;
end ;

{---------------------------------------------------------------------------------------}
function TOF_MulCorRub.RecupAxe(Nature : string ;var Axe : string;var LonAxe : integer;var NumAxe :  TFichierBase ) : boolean;
{---------------------------------------------------------------------------------------}
begin
  Result:=True ;
       if Nature='RU1' then begin Axe:='A1' ; NumAxe:=fbAxe1 ; end
  else if Nature='RU2' then begin Axe:='A2' ; NumAxe:=fbAxe2 ; end
  else if Nature='RU3' then begin Axe:='A3' ; NumAxe:=fbAxe3 ; end
  else if Nature='RU4' then begin Axe:='A4' ; NumAxe:=fbAxe4 ; end
  else begin Axe:='' ; Result:=False ; end;
  if Result then LonAxe := VH^.Cpta[NumAxe].Lg;
end ;

{TRAITEMENT DES COMPTES DE CORRESPONDANCES}
{---------------------------------------------------------------------------------------}
procedure TOF_MulCorRub.TraiteCorresp(TypeCor,Corresp,Libelle : string ) ;
{---------------------------------------------------------------------------------------}
var
  Q       : TQuery ;
  TheList : HTStringList ;

  Nom,Famille,
  TypeRub,Compte1,
  Exclus1,Axe,
  Sql,Prefix,
  Cor,
  Ch1,Table,TheMax,TheMin : string ;
begin
  Nom:=TypeCor+Corresp;
  Famille:=LaFamille.text  ;
  TypeRub:='GEN' ;
  Axe:='' ;

  if not (TypeCompte.ItemIndex in [0, 1]) then begin
    TypeRub:='ANA' ;
    Axe:=copy(TypeCor, 1, 2);
  end;

  TheList:=HTStringList.Create ;
  case TypeCompte.ItemIndex of
    0 :  begin Prefix := 'G_'; Table := 'GENERAUX'; Ch1 := Prefix + 'GENERAL'   ;end;
    1 :  begin Prefix := 'T_'; Table := 'TIERS'   ; Ch1 := Prefix + 'AUXILIAIRE';end;
    else begin Prefix := 'S_'; Table := 'SECTION' ; Ch1 := Prefix + 'SECTION'   ;end;
  end ;

  if PlanCompte.ItemIndex = 0 then Cor := Prefix + 'CORRESP1'
                              else Cor := Prefix + 'CORRESP2' ;

  Sql := 'SELECT MIN('+Ch1+'),MAX('+Ch1+') FROM '+Table+' WHERE '+Cor+'="'+Corresp+'"' ;

  if Axe <> '' then
    Sql:=Sql+' AND '+Prefix+'AXE="'+Axe+'"' ;

  Q := OpenSql(Sql,True) ;
  TheMin := Q.Fields[0].AsString ;
  TheMax := Q.Fields[1].AsString ;
  ferme(Q) ;

  if (TheMax = '') or (TheMin = '') then begin
    AfficheMsg(5,Corresp,'');
    Exit ;
  end ;

  Sql:='SELECT '+Ch1+','+Cor+' FROM '+Table+' WHERE '+Ch1+'>=("'+TheMin+'") AND '+Ch1+'<=("'+TheMax+'") ' ;
  if Axe <> '' then
    Sql := Sql + ' AND ' + Prefix + 'AXE="' + Axe + '"' ;

  Q:=OpenSql(Sql,True) ;
  while not Q.Eof do begin
    TheList.Add(Q.Fields[0].AsString+';'+Q.Fields[1].AsString+';') ;
    Q.next ;
  end ;

  Ferme(Q);

  if GenereListePlusCompacte(Corresp,TheList,Compte1,Exclus1) then
     InsertOrUpdateRubriques(Nom,Libelle,Famille,TypeRub,Compte1,Exclus1,Axe,'-') ;
  TheList.Free ;
end ;

{---------------------------------------------------------------------------------------}
function TOF_MulCorRub.GenereListePlusCompacte(Corresp: string ; Liste: HTStrings;var Compte,Exclus: string): boolean ;
{---------------------------------------------------------------------------------------}
var
  i         : Integer ;
  Okok,
  OkOkAvant : Boolean ;

  OldCompte,LeCorresp,
  LaChaine,LeCompte,
  Compte1,Compte2,
  Exclus2,Compte3,
  Exclus3   : string ;
begin
  Compte1:='' ; Compte2:='' ;Exclus2:='' ; Compte3:='' ; Exclus3:='' ;
  OkOk:=True ; OkOkAvant:=True ; LeCompte:='' ; OldCompte:='' ; LeCorresp:='' ;

  if Liste.Count = 0 then begin
    Result := False ;
    Exit ;
  end ;

  for i:=0 to Liste.Count-1 do begin
    OkokAvant := Okok ;
    LaChaine  := Liste.Strings[i] ;
    LeCompte  := ReadTokenSt(LaChaine) ;
    LeCorresp := ReadTokenSt(LaChaine) ;

    if i = 0 then
      Compte3 := LeCompte ;
    if i = Liste.Count-1 then
      Compte3 := Compte3 + LeCompte + ModeSolde ;

    if LeCorresp = Corresp then begin
      Okok:=True ;
      Compte1 := Compte1 + LeCompte + ModeSolde + ';';
    end
    else begin
      Okok := False ;
      Exclus3 := Exclus3 + LeCompte + ModeSolde + ';' ;
    end ;

    if (OkokAvant <> Okok) or (i = 0) then begin
      if Okok then begin
        Compte2 := Compte2 + LeCompte + ':';
        if i <> 0 then
          Exclus2 := Exclus2 + OldCompte + ';';
      end
      else begin
        Compte2 := Compte2 + OldCompte + ModeSolde + ';';
        if i<>0 then
          Exclus2 := Exclus2 + LeCompte + ':';
      end ;
    end ;

    OldCompte:=LeCompte ;
  end ;
  if (OkokAvant = Okok) then begin
    if Okok then Compte2 := Compte2 + OldCompte + ModeSolde + ';'
            else Exclus2 := Exclus2 + OldCompte + ';' ;
  end ;

  Result:=True ;

  if (length(Compte2)<250) and (length(exclus2)<250) then begin Compte:=Compte2 ; Exclus:=Exclus2 ; end
  else if length(Exclus3)<250 then begin Compte:=Compte3 ; Exclus:=Exclus3 ; end
  else if length(Compte1)<250 then begin Compte:=Compte1 ; Exclus:='' ; end
  else Result:=False ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_MulCorRub.BHelpClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  CallHelpTopic(Ecran) ;
end;

Initialization
  RegisterClasses([TOF_MulCorRub]);
end.

