unit UTomNomen;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HDB,HEnt1,HCrc,HMsgBox,UTOM, UTob,
      // Modif BTP
      menus,NomenUtil,saisutil,m3fp,ChoixDupNomen,
     // Objet génération des fichier XLS dans un répertoire
     {$IFDEF BTP}
      UtilsMetresXLS, olectnrs, UtilXlsBTP, FileCtrl,
     {$ENDIF}
{$IFDEF EAGLCLIENT}
      MaineAGL, eFichList,
{$ELSE}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,fichlist,
{$ENDIF}
{$IFDEF GPAOLIGHT}
      EntGC, wJetons,
{$ENDIF }
      AglInit,HTB97,UentCommun ;
Type
     TOM_NOMENENT = Class (TOM)
      private
        DomaineDef : string;
        BDuplic : TToolBarButton97;
        NomenCreation : boolean;
{$IFDEF BTP}
        // Modif BTP
        TOBDetailEntete : TOB;
        Bmenu : TToolbarButton97;
        PPopmenu : TpopupMenu;
        // -----
        // Procédures & variables propres à la gestion des métrés
        BExcel      : TToolbarButton97;
        MetreArticle: TMetreArt;
        RepDepart   : String;
        RepArrivee  : String;
        LastOuvrage : string;
        BAppelVariable: TToolbarButton97;       
{$ENDIF}
        TOBD,TOBE: TOB;
{$IFDEF GPAOLIGHT}
        TobGP : TOB;
{$ENDIF GPAOLIGHT}
{$IFDEF BTP}
        procedure ChargeTob(Article: String);
        procedure CalculeOuvrage(Article: String);
        procedure OuvreVariable(Sender: Tobject);
        procedure ExcelOnClick(Sender: Tobject);
        procedure AfficheValorisation(TOBLOC: TOB);
{$ENDIF}
        procedure Duplication(ARTICLE,OuvrageRef: String; Var CodeRes : string);
//        function AffecteCodeNomen(CodeArticle,CodeNomenRef: string): string;
//        function ExistNomen (Article,Nomen : string) : boolean;
        procedure RecopieNomenclature(Article, OuvrageRef,Ouvragedest: string);
        procedure InsertNomen;
        procedure BDuplicClick(Sender: TOBject);
        // -----
      public
        procedure OnChangeField (F : TField)  ; override ;
        procedure OnUpdateRecord  ; override ;
        procedure OnLoadRecord  ; override ;
        procedure OnNewRecord  ; override ;
        procedure OnDeleteRecord  ; override ;
        // Modif BTP
        procedure OnArgument(S: String); override;
        procedure OnClose         ; override ;
     end ;

const
	// libellés des messages
	TexteMessage: array[1..10] of string 	= (
          {1}        'Type de nomenclature obligatoire ou incorrect',
          {2}        'Code nomenclature obligatoire',
          {3}        'Libellé de la nomenclature obligatoire',
          {4}        'Suppression impossible, nomenclature utilisée dans au moins une autre',
          {5}        'Le code de nomenclature est déjà utilisé #13pour l''article composé ',
          {6}        'Veuillez renseigner la quantité d''expression du détail',
          {7}        'L''affectation automatique du code nomenclature est impossible',
          {8}        'L''affectation automatique du code ouvrage est impossible',
          {9}        'Erreur Lors de la tentative de création de l''élément',
          {10}       ''
                     );
// Modif BTP
procedure AGLCalculeOuvrage(parms:array of variant; nb: integer ) ;
// ----

implementation
uses
{$IFDEF GPAOLIGHT}
  wNomeLig,
{$ENDIF GPAOLIGHT}
  UtilFichiers,
  ParamSoc ;

/////////////////////////////////////////////
// ****** TOM NOMEN ***********************
/////////////////////////////////////////////
procedure TOM_NOMENENT.BDuplicClick (Sender : TOBject);
var CodeRes : string;
begin
duplication (THedit(GetCOntrol('GNE_ARTICLE')).Text,THedit(GetCOntrol('GNE_NOMENCLATURE')).Text,CodeRes);
if CodeRes <> '' then
   begin
   {$IFDEF BTP}
   CalculeOuvrage (CodeRes);
   {$ENDIF}
   TFFicheListe(ecran).OM.RefreshDB ;
   end;
end;

procedure TOM_NOMENENT.OnArgument(S: String);
{$IFDEF BTP}
var Critere   : String;
    i_ind     : Integer;
    ChampMul  : string;
    valmul    : string;
{$ENDIF}
begin
  inherited;
BDuplic := TToolBarButton97 (ecran.FindComponent('B_DUPLICATION'));
BDuplic.Onclick := BDuplicClick;
  {$IFDEF BTP}
  //ecran.Width := ecran.width + 160;
  TToolBarButton97(ecran.FindComponent ('Bimprimer')).Visible := false;
  {$IFDEF EAGL}
  {$ELSE}
(*
THDBGRid(GetControl('fliste')).width := THDBGRid(GetControl('fliste')).width + 150;
THDBGRid(GetControl('fliste')).left := THDBGRid(GetControl('fliste')).left - 160;
THDBGRid(GetControl('fliste')).Columns.Items [0].Width := THDBGRid(GetControl('fliste')).Columns.Items [0].Width + 50;
THDBGRid(GetControl('fliste')).Columns.Items [1].Width := THDBGRid(GetControl('fliste')).Columns.Items [1].Width + 100;
THDBGrid(GetCOntrol('fliste')).Invalidate;
THDBGrid(GetCOntrol('fliste')).Refresh;
*)
  {$ENDIF}
Bmenu := TToolbarButton97(ecran.FindComponent ('BCOMPOSANT'));
Bmenu.Hint := TraduireMemoire('Accès au Sous-détail');
PpopMenu := Tpopupmenu (ecran.FindComponent ('POPZBTP'));
Bmenu.DropdownMenu := Ppopmenu;
Ecran.Caption := TraduireMemoire('Ouvrage (Entête) :');
SetControlVisible ('GNE_QTEDUDETAIL',true);
SetControlVisible ('TGNE_QTEDUDETAIL',true);
SetControlVisible ('TUNITE',true);
ecran.Height := 420;
SetControlVisible ('GValorisation', true);
THNumEdit(GetControl('VCOEFPAPR')).Masks.PositiveMask :='#0.0000';
THNumEdit(GetControl('VCOEFPRPV')).Masks.PositiveMask :='#0.0000';
THNumedit(GetControl('VPA')).Decimals := V_PGI.okDecP;
THNumedit(GetControl('VPA')).NumericType := ntDecimal;
THNumedit(GetControl('VPR')).Decimals := V_PGI.okDecP;
THNumedit(GetControl('VPR')).NumericType := ntDecimal;
THNumedit(GetControl('VPVHT')).Decimals := V_PGI.okDecP;
THNumedit(GetControl('VPVHT')).NumericType := ntDecimal;
  Repeat
    Critere:=uppercase(Trim(ReadTokenSt(S))) ;
    if Critere<>'' then
      begin
      i_ind:=pos('=',Critere);
      if i_ind>0 then
         begin
         ChampMul:=copy(Critere,1,i_ind-1);
         ValMul:=copy(Critere,i_ind+1,length(Critere));
         if champMul = 'ARTICLE' then
            begin
            chargeTob (ValMul);
            SetControlCaption ('TUNITE',UniteOuvrage(Valmul));
            end;
         end;
      end;
  until Critere='';

  
  if TToolbarButton97(ecran.FindComponent('BTEXCEL')) <> nil then
  begin
    BExcel         := TToolbarButton97(GetControl('BTEXCEL'));
    BExcel.visible := false;
    BExcel.onclick := ExcelOnClick;
  end;

  BAppelVariable := TToolbarButton97(GetControl('BAPPELVARIABLE'));
  BAppelVariable.onclick := OuvreVariable;
{$ENDIF}

{$IFDEF GPAOLIGHT}
  if VH_GC.OASeria then  TobGP := TOB.create('WNOMETET',nil,-1);
{$ENDIF }

  TOBE := TOB.create ('NOMENENT',nil,-1);
  TOBD := TOB.create ('LESNOMENLIG',nil,-1);

end;

procedure TOM_NOMENENT.OnChangeField(F: TField);
begin
    Inherited;
end;

/////////////////////////////////////////////////////////////////////////////

procedure TOM_NOMENENT.OnUpdateRecord  ;
var QCode : TQuery;
    CodArt: string;
    ind   : Integer;
{$IFDEF GPAOLIGHT}
    TOBA, TobWAN : TOB;
{$ENDIF GPAOLIGHT}
begin
    Inherited;
  CodArt:='';

  If (GetField ('GNE_TYPNOMENC')= '') then
  begin
    SetFocusControl ('GNE_TYPNOMENC');
    LastError := 1 ;
    LastErrorMsg := TexteMessage[LastError] ;
    exit;
  end;

  If (GetField ('GNE_NOMENCLATURE')= '') then
  begin
    SetFocusControl ('GNE_NOMENCLATURE');
    LastError := 2 ;
    LastErrorMsg := TexteMessage[LastError] ;
    exit;
  end;

//
// Modif BTP
//
{$IFDEF BTP}
//If (valeur(GetField ('GNE_QTEDUDETAIL')) <= 0) then
If (valeur(GetField ('GNE_QTEDUDETAIL')) < 0) then
   begin
   SetFocusControl ('GNE_QTEDUDETAIL');
   LastError := 6 ;
   LastErrorMsg := TexteMessage[LastError] ;
   exit;
   end;
CalculeOuvrage(getfield('GNE_NOMENCLATURE'));
{$ELSE}
   SetField ('GNE_QTEDUDETAIL','1');
{$ENDIF}
//

QCode:=OpenSQL('SELECT GNE_ARTICLE FROM NOMENENT WHERE GNE_NOMENCLATURE="'
               + GetControlText('GNE_NOMENCLATURE')+'" AND GNE_ARTICLE<>"'
               + GetControlText('GNE_ARTICLE')+'"',false);
if not QCode.Eof then
   CodArt:=Copy(QCode.FindField('GNE_ARTICLE').AsString,0,18);
Ferme(QCode);
if CodArt<>'' then
   begin
   SetFocusControl ('GNE_NOMENCLATURE');
   LastError := 5;
   PGIBox(TexteMessage[LastError]+CodArt,Ecran.Caption);
   exit;
   end;

If (GetField ('GNE_LIBELLE')= '') then
    begin
    SetFocusControl ('GNE_LIBELLE');
    LastError := 3 ;
    LastErrorMsg := TexteMessage[LastError] ;
    exit;
    end;
SetField ('GNE_DATEMODIF', V_PGI.DateEntree);
if NomenCreation then SetField ('GNE_DATECREATION', V_PGI.DateEntree);
SetControlEnabled ('BCOMPOSANT', True);
{$IFDEF GPAOLIGHT}
if VH_GC.OASeria then
begin
  if (Valeur(GetField('GNE_IDENTIFIANTWNT')) = 0) then
  begin
    QCode:=OpenSQL('SELECT WNT_MAJEUR FROM WNOMETET WHERE WNT_NATURETRAVAIL="FAB" AND WNT_ARTICLE="'
                   + GetControlText('GNE_ARTICLE') + '" Order By WNT_MAJEUR Desc', True);
    if not QCode.Eof then
      CodArt := QCode.FindField('WNT_MAJEUR').AsString;
    Ferme(QCode);
    if Trim(CodArt) = '' then
      CodArt := 'A'
    else
      CodArt := Chr(Ord(CodArt[1]) + 1);
    TobGP.InitValeurs();
    TobGP.PutValue('WNT_NATURETRAVAIL' , 'FAB');
    TobGP.PutValue('WNT_ARTICLE'       , GetControlText('GNE_ARTICLE'));
    TobGP.PutValue('WNT_CODEARTICLE'   , Copy(GetControlText('GNE_ARTICLE'), 1, 18));
    TobGP.PutValue('WNT_MAJEUR'        , CodArt);
    TobGP.PutValue('WNT_IDENTIFIANT'   , wSetJeton('WNT'));
    TobGP.PutValue('WNT_LIBELLE'       , GetControlText('GNE_LIBELLE'));
    TobGP.PutValue('WNT_CODITI'        , 'AGC');
    TobGP.PutValue('WNT_ETATREV'       , 'MOD');
    TobGP.PutValue('WNT_QLOTSAIS'      , 1);
    TobGP.PutValue('WNT_QLOTSTOC'      , 1);
    TobGP.PutValue('WNT_COEFLOT'       , 1);
    TOBA := TOB.Create('ARTICLE', nil, -1);
    QCode:=OpenSQL('SELECT * FROM ARTICLE WHERE GA_ARTICLE="' + GetControlText('GNE_ARTICLE') + '"', True);
    TOBA.SelectDB('', QCode);
    Ferme(QCode);
    TobGP.PutValue('WNT_UNITELOT'      , TOBA.GetValue('GA_UNITEPROD'));
    TobGP.PutValue('WNT_QUALIFUNITESTO', TOBA.GetValue('GA_UNITEPROD'));
    TobA.PutValue('GA_QUALIFUNITESTO'  , TOBA.GetValue('GA_UNITEPROD'));
    TOBA.PutValue('GA_TENUESTOCK'      , 'X');
    TOBA.UpdateDB();
    TOBA.Free;
    TobGP.InsertDB(nil);
    SetField('GNE_IDENTIFIANTWNT', TobGP.GetValue('WNT_IDENTIFIANT'));
  end
  else
  begin
    QCode:=OpenSQL('SELECT WNT_NATURETRAVAIL, WNT_ARTICLE, WNT_MAJEUR, WNT_LIBELLE FROM WNOMETET WHERE WNT_IDENTIFIANT=' +
                   IntToStr(GetField('GNE_IDENTIFIANTWNT')), True);
    TobGP.SelectDB('', QCode);
    TobGP.PutValue('WNT_LIBELLE'       , GetControlText('GNE_LIBELLE'));
    TobGP.InsertOrUpdateDB();
  end;
  TobWAN := TOB.Create('WARTNAT', nil, -1);
  TobWAN.InitValeurs();
  TobWAN.PutValue('WAN_NATURETRAVAIL', 'FAB');
  TobWAN.PutValue('WAN_ARTICLE', GetControlText('GNE_ARTICLE'));
  TobWAN.PutValue('WAN_CODEARTICLE', Copy(GetControlText('GNE_ARTICLE'), 1, 18));
  TobWAN.PutValue('WAN_IDENTIFIANT', TobGP.GetValue('WNT_IDENTIFIANT'));
  TobWAN.PutValue('WAN_CODITI', 'AGC');
  TobWAN.PutValue('WAN_MISEENPROD', 'DEC');
  TobWAN.PutValue('WAN_WBMEMO', '-');
  TobWAN.InsertOrUpdateDB();
  TobWAN.Free;
end;
{$ENDIF GPAOLIGHT}

  //sauvegarde du fichier XLs du metre
  //Validation du fichier métré

  For ind := 0 to TobDetailEntete.detail.count -1 do
  begin
    // Gestion des métrés
    if MetreArticle <> nil then
    Begin
      MetreArticle.FermerMetreXLs;
      RepDepart := MetreArticle.fRepMetreLocal + MetreArticle.FormatArticle(TobDetailEntete.detail[Ind].GetString('GNE_NOMENCLATURE')) + '.xlsx';
      RepArrivee:= MetreArticle.fRepMetre + MetreArticle.FormatArticle(TobDetailEntete.detail[Ind].GetString('GNE_NOMENCLATURE')) + '.xlsx';
      MetreArticle.CopieLocaltoServeur(RepDepart, RepArrivee);
    end;
  end;

end;

//////////////////////////////////////////////////////////////////////

procedure TOM_NOMENENT.OnLoadRecord  ;
// Modif BTP
{$IFDEF BTP}
var TobLoc: TOB;
    critrech : array of variant;
{$ENDIF}
// -------
begin
  inHerited;

  SetControlEnabled ('BCOMPOSANT', (GetField('GNE_NOMENCLATURE')<>''));

  // Modif BTP
  {$IFDEF BTP}
  if (getfield('GNE_NOMENCLATURE') <> '') then
  begin
    setlength (critrech,1);
    critrech[0] := getfield('GNE_NOMENCLATURE');
    //
    TobLoc := TOBdetailEntete.FindFirst (['GNE_NOMENCLATURE'],critrech,true);
    if TObLoc <> nil then AfficheValorisation (TOBLOC);
    //
    lastOuvrage := getfield('GNE_NOMENCLATURE');
    // Gestion des métrés
    MetreArticle := TMetreArt.CreateArt('OUV', Getfield('GNE_NOMENCLATURE'));
    if not MetreArticle.ControleMetre then
    begin
      BExcel.visible := False;
      BAppelVariable.Visible := false;
    end
    else
    begin
      BExcel.visible  := True;
      BAppelVariable.Visible := True;
      //On copie le fichier de l'article se trouvant sur le serveur sur le poste Local
      if FileExists(MetreArticle.fRepMetre + MetreArticle.fFileName) then
      begin
        if not FileExists(MetreArticle.fRepMetreLocal + MetreArticle.fFileName) then
          CopieFichier(MetreArticle.fRepMetre + MetreArticle.fFileName, MetreArticle.fRepMetreLocal + MetreArticle.fFileName);
      end
      else
      begin
        if not FileExists(MetreArticle.fRepMetreLocal + MetreArticle.fFileName) then
          CopieFichier (MetreArticle.fFichierVide, MetreArticle.fRepMetreLocal + MetreArticle.fFileName);     //On charge un fichier vide...
      end;
    end;
    //
  end;
  {$ENDIF}

end;


procedure TOM_NOMENENT.OnNewRecord  ;
begin
    InHerited ;
// Modif BTP 06/02/2001
{$IFDEF BTP}
   AfficheValorisation (nil);
{$ENDIF}
// on initialise cette qte à 1 pour les traitements ultérieurs
SetField ('GNE_QTEDUDETAIL','1');
//
// --------
SetField ('GNE_DATECREATION', V_PGI.DateEntree);
SetField ('GNE_DOMAINE', DomaineDef);
SetControlEnabled ('GNE_TYPNOMENC', TRUE);
SetControlEnabled ('BCOMPOSANT', False);
NomenCreation := True;
end;

procedure TOM_NOMENENT.OnDeleteRecord  ;
var
    Trouve : Boolean;
{$IFDEF BTP}
    critrech : array of variant;
    TobLoc : TOB;
{$ENDIF}
begin

  Trouve := ExisteSQL('Select GNL_NOMENCLATURE from NOMENLIG where GNL_SOUSNOMEN="'+GetField('GNE_NOMENCLATURE')+'"');
  if Trouve then
  begin
    SetFocusControl ('GNE_NOMENCLATURE');
    LastError := 4 ;
    LastErrorMsg := TexteMessage[LastError] ;
    exit;
  end;

{$IFDEF BTP}
  ExecuteSQL ('DELETE FROM NOMENLIG WHERE GNL_NOMENCLATURE="'+GetField('GNE_NOMENCLATURE')+'"');
  setlength (critrech,1);
  critrech[0] := getfield('GNE_NOMENCLATURE');
  TobLoc := TOBDetailEntete.FindFirst (['GNE_NOMENCLATURE'],critrech,true);

  // FV => Gestion des métrés
  If MetreArticle <> nil then
  begin
    //Fermeture d'excel
    MetreArticle.CloseXLS;
    //Suppression du fichier dans répertoire poste de Travail
    MetreArticle.SupprimeFichierXLS(MetreArticle.frepMetreLocal + MetreArticle.FFileName);
    //suppression du fichier dans répertoire serveur
    MetreArticle.SupprimeFichierXLS(MetreArticle.frepMetre + MetreArticle.FFileName);
  end;

  if TObLoc <> nil then TobLoc.free ;
{$ENDIF}
end;

procedure TOM_NOMENENT.Onclose  ;
Var FichierArticle  : string;
    Ind             : Integer;
begin
    InHerited ;

  For ind := 0 to TobDetailEntete.detail.count -1 do
  begin
    // Gestion des métrés
    if MetreArticle <> nil then
    begin
      // Gestion des métrés
      FichierArticle := MetreArticle.fRepMetreLocal + MetreArticle.FormatArticle(TobDetailEntete.detail[Ind].GetString('GNE_NOMENCLATURE')) + '.xlsx';
      DeleteFichier(FichierArticle);
    end;
  end;
  
  FreeAndNil(MetreArticle);

  TOBDetailEntete.free;

  TOBE.free;
  TOBD.free;

  {$IFDEF GPAOLIGHT}
  if VH_GC.OASeria then TobGP.free;
  {$ENDIF GPAOLIGHT}



end;

{$IFDEF BTP}
procedure TOM_NOMENENT.ChargeTob(Article: String);
Var
Q : Tquery;
Req : String;
Indice : Integer;
valeurs : T_valeurs;
TOBLIG: TOB;
begin
Req := 'Select GA_DOMAINE from ARTICLE WHERE GA_ARTICLE="' + Article + '"';
Q := OpenSql (Req,true);
if not Q.eof then
begin
  DomaineDef := Q.findField('GA_DOMAINE').AsString;
end;
ferme (Q);

TOBDetailEntete := TOB.create ('',nil,-1);

Req := 'Select * from NOMENENT WHERE GNE_ARTICLE="' + Article + '"';
Q:=opensql (Req,true);
if not Q.eof then
begin
   TOBDetailEntete.LoadDetailDB ('NOMENENT','','',Q,true);
   for Indice := 0 to TOBdetailEntete.Detail.Count-1 do
   begin
      TOBLig := TobDetailEntete.Detail[Indice];
      TOBLig.AddChampSup ('VPA',true);
      TOBLig.AddChampSup ('VPR',true);
      TOBLig.AddChampSup ('VPVHT',true);
      TOBLig.AddChampSup ('VPVTTC',true);
      ValoriseOuvrage ( TOBLig.getvalue('GNE_NOMENCLATURE'),nil,nil,valeurs);
      FormatageTableau (valeurs);
      TOBLig.putvalue('VPVHT',valeur(Strs(Valeurs[2],V_PGI.OkDecP)));
      TOBLig.putvalue('VPA',valeur(Strs(Valeurs[0],V_PGI.OkDecP)));
      TOBLig.putvalue('VPR',valeur(Strs(Valeurs[1],V_PGI.OkDecP)));
      TOBLig.putvalue('VPVTTC',valeur(Strs(Valeurs[3],V_PGI.OkDecP)));
   end;
end;
ferme (Q);
end;
{$ENDIF}
/////////////////////////////////////////////////////////////////////////////

procedure AGLCalculeOuvrage(parms:array of variant; nb: integer ) ;
{$IFDEF BTP}
var  F : TForm ;
     OM : TOM ;
{$ENDIF}
begin
{$IFDEF BTP}
F:=TForm(Longint(Parms[0])) ;
if (F is TFFicheListe) then OM:=TFFicheListe(F).OM else exit;
if (OM is TOM_NOMENENT) then TOM_NOMENENT(OM).CalculeOuvrage(Parms[1]) else exit;
{$ENDIF}
end;

{$IFDEF BTP}
procedure TOM_NOMENENT.CalculeOuvrage(Article: String);
var
   critrech : array of variant;
   TOBLOc : TOB;
   Valeurs : T_Valeurs;
begin
   setlength (critrech,1);
   critrech[0] := Article;
   // Mise à jour de la tob des montants
   TobLoc := TOBdetailEntete.FindFirst (['GNE_NOMENCLATURE'],critrech,true);
   if TObLoc <> nil then
   begin
      ValoriseOuvrage ( TOBLoc.getvalue('GNE_NOMENCLATURE'),nil,nil,valeurs);
      TOBLoc.putvalue('VPVHT',valeur(Strs(Valeurs[2],V_PGI.OkDecV)));
      TOBLoc.putvalue('VPA',valeur(Strs(Valeurs[0],V_PGI.OkDecV)));
      TOBLoc.putvalue('VPR',valeur(Strs(Valeurs[1],V_PGI.OkDecV)));
      TOBLoc.putvalue('VPVTTC',valeur(Strs(Valeurs[3],V_PGI.OkDecV)));
   end
   else
   begin
      TOBLoc := TOB.create ('NOMENENT',TobdetailEntete,-1);
      TOBLoc.AddChampSup ('VPA',false);
      TOBLoc.AddChampSup ('VPR',false);
      TOBLoc.AddChampSup ('VPVHT',false);
      TOBLoc.AddChampSup ('VPVTTC',false);
      TOBLoc.putvalue('GNE_NOMENCLATURE',getfield('GNE_NOMENCLATURE'));
      TOBLoc.putvalue('VPVHT',valeur('0'));
      TOBLoc.putvalue('VPA',valeur('0'));
      TOBLoc.putvalue('VPR',valeur('0'));
      TOBLoc.putvalue('VPVTTC',valeur('0'));
   end;
   AfficheValorisation (TOBLOC);
end;
{$ENDIF}

{$IFDEF BTP}
procedure TOM_NOMENENT.AfficheValorisation (TOBLOC : TOB);
var
   DValeur : double;
begin
   if TOBLOC <> nil then
   begin
      SetControlProperty('VPA','Value',TobLoc.getvalue('VPA'));
      SetControlProperty('VPR','Value',TobLoc.getvalue('VPR'));
      SetControlProperty('VPVHT','Value',TobLoc.getvalue('VPVHT'));
      SetControlProperty('VPVTTC','Value',TobLoc.getvalue('VPVTTC'));
      if (TobLoc.getvalue('VPA') <> valeur ('0')) then
         Dvaleur := TobLoc.getvalue('VPR')/TobLoc.getvalue('VPA')
      else
         DValeur := valeur('0');
      SetControlProperty('VCOEFPAPR','Value',Dvaleur);
      if TobLoc.getvalue('VPR') <> valeur ('0') then
         Dvaleur := TobLoc.getvalue('VPVHT')/TobLoc.getvalue('VPR')
      else
         DValeur := valeur('0');
      SetControlProperty('VCOEFPRPV','Value',Dvaleur);
   end
   else
   begin
      SetControlProperty('VPA','Value',valeur('0'));
      SetControlProperty('VPR','Value',valeur('0'));
      SetControlProperty('VPVHT','Value',valeur('0'));
      SetControlProperty('VPVTTC','Value',valeur('0'));
      SetControlProperty('VCOEFPAPR','Value',valeur('0'));
      SetControlProperty('VCOEFPRPV','Value',valeur('0'));
   end;
end;
{$ENDIF}

/////////////////////////////////////////////////////////////////////////////

procedure TOM_NOMENENT.Duplication(ARTICLE,OuvrageRef: String; var CodeRes : String);
var TOBResult : TOB;
    CodeNomenRet : string;
    Nouveaufichier : string;
begin
  TOBResult := TOB.create ('',nil,-1);
  TOBResult.addchampsupValeur('ARTICLE',Article,false);
  if ChoixDuplication (TOBResult) then
   begin
   if TOBResult.GetValue('AUTOMATIQUE') = 'X' then
      begin
      CodeNomenRet := AffecteCodeNomen(Article,OuvrageRef);
      if CodeNomenRet = '' then
         begin
{$IFDEF BTP}
         LastError := 8 ;
         LastErrorMsg := TexteMessage[LastError] ;
{$ELSE}
         LastError := 7 ;
         LastErrorMsg := TexteMessage[LastError] ;
{$ENDIF}
         TOBResult.free;
         exit;
         end ;
         TOBResult.putValue('CODENOMEN',AffecteCodeNomen(Article,OuvrageRef));
      end;
   RecopieNomenclature(Article,OuvrageRef,TobResult.GetValue('CODENOMEN'));
   CodeRes := TOBResult.GetValue('CODENOMEN');
   end;

   TOBResult.free;

{$IFDEF BTP}
  //Ici il faudrait dupliquer le fichier métré de départ avec le nouveau code article...
  if MetreArticle <> nil then
  begin
    if CodeRes <> '' then
    begin
      MetreArticle.DuplicationFichierXls(MetreArticle.fFileName, CodeRes + '.xlsx');
    end;
  end;
{$ENDIF}

end;

procedure TOM_NOMENENT.InsertNomen;
begin
if Not TOBE.InsertDB(Nil,true) then V_PGI.IoError:=oeUnknown ;
if V_PGI.IOError = OeOk then
   if not TOBD.InsertDBByNivel(true) then V_PGI.IoError:=oeUnknown ;
{$IFDEF GPAOLIGHT}
if VH_GC.OASeria then
begin
    if V_PGI.IOError = OeOk then
       if not TobGP.InsertDBByNivel(true) then V_PGI.IoError:=oeUnknown ;
end;
{$ENDIF GPAOLIGHT}
end;

procedure TOM_NOMENENT.RecopieNomenclature(Article,OuvrageRef,Ouvragedest : string);
var TOBL,TOBR : TOB;
    QQ : Tquery;
    Indice : integer;
{$IFDEF GPAOLIGHT}
    TOBA,TobLGP, TobWAN : TOB;
    WNTMajeur : string;
{$ENDIF GPAOLIGHT}

begin
TOBR := TOB.create ('NOMENENT',nil,-1);
QQ := opensql('Select * from NOMENENT where GNE_ARTICLE="'+Article+'" AND GNE_NOMENCLATURE="'+OuvrageRef+'"',false);
TOBR.SelectDB ('',QQ);
ferme(QQ);
// Entete
TOBE.Dupliquer (TOBR,true,true);
TOBR.Free;
{$IFDEF GPAOLIGHT}
if VH_GC.OASeria then
begin
    QQ:=OpenSQL('SELECT WNT_MAJEUR FROM WNOMETET WHERE WNT_NATURETRAVAIL="FAB" AND WNT_ARTICLE="' +
                Article + '" Order By WNT_MAJEUR Desc', True);
    if not QQ.Eof then
       WNTMajeur := QQ.FindField('WNT_MAJEUR').AsString;
    Ferme(QQ);
    if Trim(WNTMajeur) = '' then
        WNTMajeur := 'A'
        else
        WNTMajeur := Chr(Ord(WNTMajeur[1]) + 1);
    TobGP.InitValeurs();
    TobGP.PutValue('WNT_NATURETRAVAIL', 'FAB');
    TobGP.PutValue('WNT_ARTICLE'      , Article);
    TobGP.PutValue('WNT_CODEARTICLE'  , Copy(Article, 1, 18));
    TobGP.PutValue('WNT_MAJEUR'       , WNTMajeur);
    TobGP.PutValue('WNT_IDENTIFIANT'  , wSetJeton('WNT'));
    TobGP.PutValue('WNT_CODITI'       , 'AGC');
    TobGP.PutValue('WNT_ETATREV'      , 'MOD');
    TobGP.PutValue('WNT_QLOTSAIS'     , 1);
    TobGP.PutValue('WNT_QLOTSTOC'     , 1);
    TobGP.PutValue('WNT_COEFLOT'      , 1);
    TOBA := TOB.Create('ARTICLE', nil, -1);
    QQ:=OpenSQL('SELECT * FROM ARTICLE WHERE GA_ARTICLE="' + Article + '"', True);
    TOBA.SelectDB('', QQ);
    Ferme(QQ);
    TobGP.PutValue('WNT_UNITELOT'      , TOBA.GetValue('GA_UNITEPROD'));
    TobGP.PutValue('WNT_QUALIFUNITESTO', TOBA.GetValue('GA_UNITEPROD'));
    TOBE.PutValue('GNE_IDENTIFIANTWNT' , TobGP.GetValue('WNT_IDENTIFIANT'));
    TobWAN := TOB.Create('WARTNAT', nil, -1);
    TobWAN.InitValeurs();
    TobWAN.PutValue('WAN_NATURETRAVAIL', 'FAB');
    TobWAN.PutValue('WAN_ARTICLE', Article);
    TobWAN.PutValue('WAN_CODEARTICLE', Copy(Article, 1, 18));
    TobWAN.PutValue('WAN_IDENTIFIANT', TobGP.GetValue('WNT_IDENTIFIANT'));
    TobWAN.PutValue('WAN_CODITI', 'AGC');
    TobWAN.PutValue('WAN_MISEENPROD', 'DEC');
    TobWAN.PutValue('WAN_WBMEMO', '-');
    TobWAN.InsertOrUpdateDB();
    TobWAN.Free;
    SetField('GNE_IDENTIFIANTWNT', TobGP.GetValue('WNT_IDENTIFIANT'));
    TOBA.Free;
end;
{$ENDIF GPAOLIGHT}
TOBE.PutValue('GNE_NOMENCLATURE',OuvrageDest);
TOBE.SetAllModifie(true);
// detail
TOBR := TOB.Create ('LENOMENLIG',nil,-1);
QQ := opensql('Select * from NOMENLIG where GNL_NOMENCLATURE="'+OuvrageRef+'"',false);
TOBR.LoadDetailDB ('NOMENLIG','','',QQ,false,true);
ferme (QQ);
TOBD.Dupliquer (TOBR,true,true);
TOBR.free;
for Indice:=0 to TOBD.detail.count -1 do
    begin
    TOBL := TOBD.Detail[indice];
    TOBL.PutValue('GNL_NOMENCLATURE',OuvrageDest);
    TOBD.SetAllModifie(true);
{$IFDEF GPAOLIGHT}
    if VH_GC.OASeria then
    begin
        TobLGP := TOB.Create('WNL', TobGP, -1);

        { Champs }
        TobLGP.AddChampSupValeur('WNL_COMPOSANT'    , TOBL.GetValue('GNL_ARTICLE'));
        TobLGP.AddChampSupValeur('WNL_CODECOMPOSANT', Copy(TobLGP.GetValue('WNL_COMPOSANT'), 1, 18));
        TobLGP.AddChampSupValeur('WNL_QLIENSAIS'    , TOBL.GetValue('GNL_QTE'));
        TobLGP.AddChampSupValeur('WNL_PHASE'        , 'AGC');
        TobLGP.AddChampSupValeur('WNL_OPEITI'       , '110');
        TobLGP.AddChampSupValeur('WNL_OPEITIAPPRO'  , '110');
        TobLGP.AddChampSupValeur('WNL_MODECONSO'    , 'LAN');
        TobLGP.AddChampSupValeur('WNL_QLOTSAIS', 1.0);
        
(*        TobLGP.InitValeurs;
        TobLGP.PutValue('WNL_NATURETRAVAIL', 'FAB');
        TobLGP.PutValue('WNL_ARTICLE'      , TobGP.GetValue('WNT_ARTICLE'));
        TobLGP.PutValue('WNL_CODEARTICLE'  , Copy(TobGP.GetValue('WNT_ARTICLE'), 1, 18));
        TobLGP.PutValue('WNL_MAJEUR'       , TobGP.GetValue('WNT_MAJEUR'));
        TobLGP.PutValue('WNL_LIENNOME'     , ((Indice + 1) * 1000));
        TobLGP.PutValue('WNL_CIRCUIT'      , '');
        TobLGP.PutValue('WNL_IDENTIFIANT'  , wSetJeton('WNL'));
        TobLGP.PutValue('WNL_PHASE'        , 'AGC');
        TobLGP.PutValue('WNL_OPEITI'       , '110');
        TobLGP.PutValue('WNL_OPEITIAPPRO'  , '110');
        TobLGP.PutValue('WNL_COMPOSANT'    , TOBL.GetValue('GNL_ARTICLE'));
        TobLGP.PutValue('WNL_CODECOMPOSANT', Copy(TobLGP.GetValue('WNL_COMPOSANT'), 1, 18));
        TobLGP.PutValue('WNL_QLIENSAIS'    , TOBL.GetValue('GNL_QTE'));
        TobLGP.PutValue('WNL_QLIENSTOC'    , TOBL.GetValue('GNL_QTE'));
        TobLGP.PutValue('WNL_QLOTSAIS'     , 1);
        TobLGP.PutValue('WNL_MODECONSO'    , 'LAN');
        TOBA := TOB.Create('ARTICLE', nil, -1);
        QQ:=OpenSQL('SELECT * FROM ARTICLE WHERE GA_ARTICLE="'
                   + TOBL.GetValue('GNL_ARTICLE') + '"', True);
        TOBA.SelectDB('', QQ);
        Ferme(QQ);
        TobLGP.PutValue('WNL_UNITELOT'      , TOBA.GetValue('GA_UNITEPROD'));
        TobLGP.PutValue('WNL_UNITELIEN'     , TOBA.GetValue('GA_UNITEPROD'));
        TobLGP.PutValue('WNL_QUALIFUNITESTO', TOBA.GetValue('GA_UNITEPROD'));
        TobLGP.PutValue('WNL_UNITEPPER'     , TOBA.GetValue('GA_UNITEPROD'));
        TobLGP.PutValue('WNL_UNITEPERIODE'  , TOBA.GetValue('GA_UNITEPROD'));
        TobLGP.PutValue('WNL_UNITEPFIXE'    , TOBA.GetValue('GA_UNITEPROD'));
        TOBA.Free; *)
    end;
{$ENDIF GPAOLIGHT}
    end;

{$IFDEF GPAOLIGHT}
wCreateWNL(TobGP);
{$ENDIF GPAOLIGHT}

if TRANSACTIONS (INSERTNOMEN,1) <> oeOk then MessageAlerte(TexteMessage[9]) ;
TOBE.ClearDetail;
TOBD.ClearDetail;
{$IFDEF GPAOLIGHT}
if VH_GC.OASeria then
    TobGP.ClearDetail;
{$ENDIF GPAOLIGHT}
end;

(*
function TOM_NOMENENT.ExistNomen (Article,Nomen : string) : boolean;
begin
result := ExisteSQL('Select GNE_NOMENCLATURE from NOMENENT where GNE_NOMENCLATURE="'+Nomen+'"');
end;

function TOM_NOMENENT.AffecteCodeNomen (CodeArticle,CodeNomenRef : string) : string;
var Indice : Integer;
    NewCode : String;
    ok : boolean;
begin
Indice := 1;
ok := false;
result := '';
if length(trim(CodeNomenRef)) < 35 then
   begin
   repeat
   NewCode := trim(CodeNomenRef)+inttostr(Indice);
   if length(NewCode) > 35 then break;
   ok := ExistNomen (CodeArticle,NewCode);
   if ok then inc(Indice) else result := NewCode;
   until not ok;
   end else
   begin
   repeat
   if length(inttostr(Indice))> 10 then break;
   NewCode := copy (CodeNomenRef,1,35-length(inttostr(Indice)))+inttostr(Indice);
   ok := ExistNomen (CodeArticle,NewCode);
   if ok then inc(Indice) else result := NewCode;
   until not ok;
   end;
end;
*)

{$IFDEF BTP}
procedure TOM_NOMENENT.ExcelOnClick(Sender: Tobject);
begin

  if not (DS.State in [dsInsert, dsEdit]) then DS.edit; // pour passer DS.state en mode dsEdit

  MetreArticle.OuvrirMetreXLs;

end;
//
procedure TOM_NOMENENT.OuvreVariable(Sender: Tobject);
begin

  AGLLanceFiche('BTP', 'BTVARIABLE', 'B;'+Getfield ('GNE_NOMENCLATURE'),'', 'LIBELLEARTICLE=' + Getfield('GNE_LIBELLE') + ';TYPEARTICLE=' + 'OUV');

end;
{$ENDIF}

Initialization
registerclasses([TOM_NOMENENT]) ;
RegisterAglProc('CalculeOuvrage',TRUE,1,AglCalculeOuvrage);
end.
