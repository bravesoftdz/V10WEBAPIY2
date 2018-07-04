{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 06/09/2016
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTIMPORTMULDEV ()
Mots clefs ... : TOF;BTIMPORTMULDEV
*****************************************************************}
Unit BTIMPORTMULDEV_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,Fe_Main,
{$else}
     eMul,
{$ENDIF}
     uTob, 
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     HMsgBox,
     HTB97,
     UtilsRapport,
     UTOF,
     AglInit ;

Type
  TOF_BTIMPORTMULDEV = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    RapportGen   : TGestionRapport;
    fTOBresult : TOB;
    procedure BouvrirClick (Sender : Tobject);
    procedure TraiteDevisAtt (ID : String);
    function LoadlesDatas (ID : String; TODDETATT,TOBDATAATT : TOB) : boolean;
    procedure AddReport (Status : string; LibelleAssocie : string;TOBINTER : TOB);
    function CreeIntervention (TOBAA : TOB) : boolean;
    procedure AddChampsSupAppel(TOBAA: TOB);
    function GetTiers(TOBT: TOB; CodeTiers: string): boolean;
    function CreeAppel (TOBAA,TOBA,TOBT : TOB): boolean;
    procedure RechercheOuCreeContact (TOBAA,TOBT,TOBINTER : TOB) ;
    function EcritContact (TOBINTER,TOBT : TOB) : boolean;
    function FindNumContact(TOBT: TOB): integer;
    function FindNumAdresse: integer;
    procedure InitPieceDevis(TOBDEVIS, TOBA, TOBAA: TOB);
    function ConstitueDevis (TOBA,TOBAA,TOBT : TOB) : boolean;
    function PositionneTraite (TOBAA : TOB) : boolean;
    function InsereDatas (TOBA,TOBAA,TOBT : TOB) : boolean;
  end ;

Implementation
uses ParamSoc,AppelsUtil,UtilPgi,ENT1,GerePiece;

procedure TOF_BTIMPORTMULDEV.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPORTMULDEV.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPORTMULDEV.OnUpdate ;
begin
  Inherited ;
  TToolbarButton97(getControl('Bouvrir')).onclick := BouvrirClick;
end ;

procedure TOF_BTIMPORTMULDEV.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPORTMULDEV.OnArgument (S : String ) ;
begin
  Inherited ;
  RapportGen   := TGestionRapport.Create(TFMul(Ecran));
  with RapportGen do
  begin
    Titre   := 'Rapport d''intégration des devis';
    Affiche := True;
    Close   := True;
    Sauve   := False;
    Print   := False;
    InitRapport;
    Visible := false;
    PosLeft := Round(TFMul(Ecran).Width / 1.5);
    PosTop  := Round(TFMul(Ecran).Top);
  end;
  fTOBresult := TOB.Create ('LA PIECE GENEREE',nil,-1);
end ;

procedure TOF_BTIMPORTMULDEV.OnClose ;
begin
  fTOBresult.free;
  Inherited ;
end ;

procedure TOF_BTIMPORTMULDEV.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPORTMULDEV.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPORTMULDEV.BouvrirClick(Sender: Tobject);
var ID : String;
    Q : Tquery;
    i : integer;
begin
  if (TFMul(Ecran).FListe=nil) then exit;
  if (not TFMul(Ecran).Fliste.AllSelected) and (TFMul(Ecran).Fliste.nbSelected=0) then exit;
  if (PGIAsk ('Vous allez intégrer les devis sélectionnés.#13#10 Confirmez-vous ?' , Ecran.Caption)<>mrYes) then exit;
  if THCheckBox(GetControl('BAA_TRAITE')).Checked then
  begin
    if (PGIAsk ('ATTENTION : Ces devis ont déjà été intégrés.#13#10 Confirmez-vous le traitement ?' , Ecran.Caption)<>mrYes) then exit;
  end;
  TRY
    if TFMul(Ecran).Fliste.AllSelected then
    BEGIN
      Q:=TFMul(Ecran).Q;
      Q.First;
      while Not Q.EOF do
      BEGIN
        ID:=Q.FindField('BAA_GUID').AsString;
        TraiteDevisAtt (ID);
        Q.NEXT;
      END;
      TFMul(Ecran).Fliste.AllSelected:=False;
    END ELSE
    BEGIN
      for i:=0 to TFMul(Ecran).Fliste.nbSelected-1 do
      begin
        TFMul(Ecran).Fliste.GotoLeBookmark(i);
        ID:=TFMul(Ecran).Fliste.datasource.dataset.FindField('BAA_GUID').AsString;
        TraiteDevisAtt (ID);
      end;
    END;
  FINALLY
    if RapportGen.Memo.Lines.count > 0 then
    begin
      RapportGen.AfficheRapport;    //affichage du rapport d'intégration
      RapportGen.Affiche := True;
    end;
  END;
  TToolBarButton97(GetControl('BCherche')).Click;
end;


procedure TOF_BTIMPORTMULDEV.AddChampsSupAppel (TOBAA : TOB);
begin
  TOBAA.AddChampSupValeur('SAISIEOK','-');
  TOBAA.AddChampSupValeur('AFFAIRE',''); // mémorisation du numéro de l'intervention.
  TOBAA.AddChampSupValeur('DEVIS',0); // mémorisation du numéro de Devis.
  TOBAA.AddChampSupValeur('AFF_APPEL0','');
  TOBAA.AddChampSupValeur('AFF_APPEL1','');
  TOBAA.AddChampSupValeur('AFF_APPEL2','');
  TOBAA.AddChampSupValeur('AFF_APPEL3','');
  TOBAA.AddChampSupValeur('AFF_AVENANT','');
end;

procedure TOF_BTIMPORTMULDEV.TraiteDevisAtt(ID: String);
var TOBAA,TOBDATAATT,TOBDETATT : TOB;
    QQ : TQuery;
    OkTrait : boolean;
    StSql : string;
    //IOERR : TIOErr;
begin
  OkTrait := false;
  TOBAA := TOB.Create ('BAFFAIREATT',nil,-1);
  TOBDATAATT := TOB.Create ('BDATAATT',nil,-1);
  TOBDETATT := TOB.Create ('BDETAILATT',nil,-1);
  TRY
    StSql := 'SELECT * FROM BAFFAIREATT WHERE BAA_GUID="'+ID+'"';
    QQ := OpenSql(StSql,true,1,'',true);
    if not QQ.eof then
    begin
      TOBAA.SelectDB('',QQ);
      AddChampsSupAppel (TOBAA);
      OkTrait := true;
    end;
    ferme (QQ);
    if OkTrait then
    begin
      OkTrait := LoadlesDatas (ID,TOBDETATT,TOBDATAATT);
      if OkTrait then
      begin
        OkTrait := CreeIntervention (TOBAA);
        if OkTrait then AddReport ('OK','',TOBAA)
      end else
      begin
        // enregistrement du fait qu'il n'y ai pas de données associés à l'intervention !!
        AddReport ('NOK','Récupération des données de l''intervention',TOBAA);
      end;
    end;
  FINALLY
    TOBAA.free;
    TOBDATAATT.free;
    TOBDETATT.free;
  END;
end;

function TOF_BTIMPORTMULDEV.LoadlesDatas(ID : String; TODDETATT,TOBDATAATT: TOB): boolean;
var QQ : TQuery;
begin
  result := true;
  QQ := OpenSQl ('SELECT * FROM BDATAATT WHERE BIA_GUID="'+ID+'"',true,-1,'',true);
  if not QQ.eof then
  begin
    TOBDATAATT.LoadDetailDB('BDATAATT','','',QQ,false);
  end;
  ferme (QQ);
  // ----
  QQ := OpenSQl ('SELECT * FROM BDETAILATT WHERE BAB_GUID="'+ID+'"',true,-1,'',true);
  if not QQ.eof then
  begin
    TODDETATT.LoadDetailDB('BDETAILATT','','',QQ,false);
  end;
  ferme (QQ);
  if (TOBDATAATT.detail.count =0) and (TODDETATT.detail.count =0) then result := false;
end;

procedure TOF_BTIMPORTMULDEV.AddReport(Status, LibelleAssocie: string;TOBINTER: TOB);
var Txt : string;
begin
  if Status = 'OK' then
  begin
    Txt := 'Génération de l''intervention '+TOBInter.getString('AFFAIRE')+' et du devis N° '+InttoStr(TOBInter.getInteger('DEVIS'));
  end else
  begin
    Txt := 'Erreur : '+libelleAssocie;
  end;
  RapportGen.SauveLigMemo(TXT);
end;

function TOF_BTIMPORTMULDEV.CreeIntervention(TOBAA: TOB): boolean;
var TOBA,TOBT : TOB;
    Part0,Part1,Part2,Part3,avenant : string;
begin
  result := false;
  TOBA := TOB.Create ('AFFAIRE',nil,-1);
  TOBT := TOB.Create ('TIERS',nil,-1);
  TRY
    if not GetTiers (TOBT,TOBAA.getString('BAA_TIERS')) then
    begin
      AddReport ('NOK','Tiers inexistant',TOBAA);
      exit;
    end;
    Part0 := 'W';
    Avenant := '00';
    Part1 := ''; Part2 := ''; Part3 := '';
    PrePrepositionneNewCodeAppel (Part1,Part2,Part3);
    if GetParamSocSecur('SO_SAISIECODEAPPEL', False) then
    begin
      TOBAA.SetString('AFF_APPEL0',Part0);
      TOBAA.SetString('AFF_APPEL1',Part1);
      TOBAA.SetString('AFF_APPEL2',Part2);
      TOBAA.SetString('AFF_APPEL3',Part3);
      TOBAA.SetString('AFF_AVENANT',Avenant);
      TheTOB := TOBAA;
      AGLLanceFiche('BTP','BTSAISIEAPP','','','');
      result := (TOBAA.GetString('SAISIEOK')='X');
      if result then
      begin
        Part0 :=TOBAA.GetString('AFF_APPEL0');
        Part1 :=TOBAA.GetString('AFF_APPEL1');
        Part2 :=TOBAA.GetString('AFF_APPEL2');
        Part3 :=TOBAA.GetString('AFF_APPEL3');
        Avenant :=TOBAA.GetString('AFF_AVENANT');
      end;
    end else result := true;
    if result then
    BEGIN
      FinaliseNewCodeAppel(Part1,Part2,Part3);
      TOBAA.setString('AFFAIRE',RegroupePartiesAppel (Part0,Part1,Part2,Part3,avenant));
      TOBA.SetString('AFF_AFFAIRE',TOBAA.getString('AFFAIRE'));
      TOBA.SetString('AFF_AFFAIRE0',Part0);
      TOBA.SetString('AFF_AFFAIRE1',Part1);
      TOBA.SetString('AFF_AFFAIRE2',Part2);
      TOBA.SetString('AFF_AFFAIRE3',Part3);
      TOBA.SetString('AFF_AVENANT',Avenant);
      BEGINTRANS;
      result := CreeAppel (TOBAA,TOBA,TOBT);
      if result then result := ConstitueDevis (TOBA,TOBAA,TOBT);
      if result then result := InsereDatas (TOBA,TOBAA,TOBT);
      if result Then result := PositionneTraite (TOBAA);
      if result then COMMITTRANS else ROLLBACK;
    END;
    if not result then  AddReport ('NOK','Création de l''intervention',TOBAA);
  FINALLY
    TOBA.free;
    TOBT.Free;
  END;
end;

function TOF_BTIMPORTMULDEV.GetTiers(TOBT: TOB; CodeTiers: string) : boolean;
var QQ : TQuery;
begin
  result := false;
  QQ := OpenSql ('SELECT * FROM TIERS WHERE T_TIERS ="'+CodeTiers+'" AND T_NATUREAUXI IN ("CLI","PRO")',true,1,'',true);
  if not QQ.eof then
  begin
    result := true;
    TOBT.SelectDb ('',QQ);
  end;
  ferme (QQ);
end;

function TOF_BTIMPORTMULDEV.ConstitueDevis (TOBA,TOBAA,TOBT : TOB) : boolean;

  procedure addlesSupligne (OneLigne : TOB);
  begin
    OneLigne.AddChampSupValeur ('TYPELIGNE','');
    OneLigne.AddChampSupValeur('LIBELLE','');
    OneLigne.AddChampSupValeur('CODEARTICLE','');
    OneLigne.AddChampSupValeur('ARTICLE','');
    OneLigne.AddChampSupValeur('QTEFACT',0);
    OneLigne.AddChampSupValeur('QUALIFQTE','');
    OneLigne.AddChampSupValeur('AVECPRIX','X');
    OneLigne.AddChampSupValeur('PUHTDEV',0);
    OneLigne.AddChampSupValeur('CODETVA','');
    OneLigne.AddChampSupValeur('LIBELLE','');
    OneLigne.AddChampSupValeur('FACTURABLE','');
    OneLigne.AddChampSupValeur('BLOCNOTE','');
  end;

  function RecupCodeTva(TOBT: TOB;taux : double) : string;
  var QQ : TQuery;
      SQl : string;
  begin
    result := '';
    SQl := 'SELECT TV_CODETAUX '+
           'FROM TXCPTTVA '+
           'WHERE '+
           'TV_TVAOUTPF="TX1" AND '+
           'TV_REGIME="'+TOBT.GetString('T_REGIMETVA')+'" AND '+
           'TV_SOCIETE="'+V_PGI.CodeSociete+'" AND '+
           'TV_TAUXVTE='+FloatToStr(Taux);

    QQ := OpenSql (Sql,true,1,'',true);
    if not QQ.eof then
    begin
      result := QQ.fields[0].AsString;
    end;
    ferme (QQ);
  end;

  function FindArticle (CodeArticle : string) : string;
  var QQ : Tquery;
  begin
    result := '';
    QQ := OpenSql ('SELECT GA_ARTICLE FROM ARTICLE WHERE GA_CODEARTICLE="'+CodeArticle+'"', true,1,'',true);
    if not QQ.eof then
    begin
      result := QQ.fields[0].AsString;
    end;
    ferme (QQ);
  end;

  function ConstitueDetailDevis (TOBDEt,TOBDEVIS,TOBA,TOBAA,TOBT : TOB) : boolean;
  var II : integer;
      OneLigne,TOBL : TOB;
      PuHT : double;
      Article,CodeTva : string;
  begin
    result := true;
    for II := 0 to TOBDet.detail.count - 1 do
    begin
      TOBL := TOBDet.detail[II];
      OneLigne := TOB.Create ('UNE LIGNE',TOBDEVIS,-1);
      addlesSupligne (OneLigne);
      Article := FindArticle(TOBL.getString('BAB_CODEARTICLE'));
      if Article = '' then
      begin
        Article := GetParamsocSecur('SO_BTARTICLEDIV','');
      end;
      if Article = '' then
      begin
        result := false;
        AddReport ('NOK','Code Article '+TOBL.getString('BAB_CODEARTICLE')+' Inconnu',TOBAA);
        break;
      end;
      OneLigne.SetString('TYPELIGNE',TOBL.getString('BAB_TYPELIGNE'));
      OneLigne.SetString('LIBELLE',TOBL.getString('BAB_LIBELLE'));
      if TOBL.getString('BAB_TYPELIGNE') <> 'COM' then
      begin
        CodeTva := RecupCodeTva(TOBT,TOBL.getDouble('BAB_TAUXTAXE'));
        if CodeTva = '' then
        begin
          result := false;
          AddReport ('NOK','TVA '+FloatToStr(TOBL.getDouble('BAB_TAUXTAXE'))+' non définie',TOBAA);
          break;
        end;
        if TOBL.getDouble('BAB_QTEFACT') <> 0 then PuHt := Arrondi(TOBL.getDouble('BAB_MONTANTHTDEV')/TOBL.getDouble('BAB_QTEFACT'),V_PGI.okdecP)
                                              else PuHt := TOBL.getDouble('BAB_MONTANTHTDEV');
        OneLigne.SetString('CODEARTICLE',TOBL.getString('BAB_CODEARTICLE'));
        OneLigne.SetString('ARTICLE',Article);
        OneLigne.SetString('QTEFACT',TOBL.getString('BAB_QTEFACT'));
        OneLigne.SetString('QUALIFQTE',TOBL.getString('BAB_QUALIFQTEVTE'));
        OneLigne.SetDouble('PUHTDEV',PuHt);
        OneLigne.SetString('CODETVA',CodeTva);
      end;
      OneLigne.SetString('LIBELLE',TOBL.getString('BAB_LIBELLE'));
      OneLigne.SetString('FACTURABLE',TOBL.getString('BAB_FACTURABLE'));
      OneLigne.SetString('BLOCNOTE',TOBL.getString('BAB_BLOCNOTE'));
    end;
  end;

var QQ : TQuery;
    TOBDEVIS,TOBDET : TOB;
begin
  result := false;
  TOBDet := TOB.Create ('LES LIGNES DET',nil,-1);
  TOBDEVIS := TOB.Create ('UN DEVIS',nil,-1);
  fTOBresult.clearDetail;
  TRY
    QQ := OpenSql ('SELECT  * FROM BDETAILATT WHERE BAB_GUID="'+TOBAA.geTString('BAA_GUID')+'" ORDER BY BAB_NUMLIGNE',true,-1,'',true);
    if not QQ.eof then
    begin
      TOBDet.LoadDetailDB('BDETAILATT','','',QQ,false);
    end;
    ferme (QQ);
    if TOBDet.detail.count > 0 then
    begin
      InitPieceDevis (TOBDEVIS,TOBA,TOBAA);
      result := ConstitueDetailDevis (TOBDEt,TOBDEVIS,TOBA,TOBAA,TOBT);
      if not result then exit;
      Result := CreatePieceFromTOB (TOBDEVIS,nil,nil,nil,'',fTOBresult);
      if result then TOBAA.SetInteger('DEVIS',fTOBresult.detail[0].getInteger('GP_NUMERO'));
    end;
  FINALLY
    TOBDet.free;
    TOBDEVIS.free;
  END;

end;

procedure TOF_BTIMPORTMULDEV.InitPieceDevis (TOBDEVIS,TOBA,TOBAA : TOB);
var CodeEtab : string;
begin
  if VH^.EtablisDefaut <> '' then CodeEtab := VH^.EtablisDefaut;
  if VH^.EtablisCpta then
  begin
    if VH^.ProfilUserC[prEtablissement].Etablissement <> '' then CodeEtab := VH^.ProfilUserC[prEtablissement].Etablissement;
  end;
  TOBDEVIS.AddChampSupValeur ('NATUREPIECEG', 'DAP');
  TOBDEVIS.AddChampSupValeur ('AFFAIRE', TOBA.getString('AFF_AFFAIRE'));
  TOBDEVIS.AddChampSupValeur ('TIERS', TOBA.getString('AFF_TIERS'));
  TOBDEVIS.AddChampSupValeur ('ETABLISSEMENT',CodeEtab );
  TOBDEVIS.AddChampSupValeur ('DOMAINE', '');
  TOBDEVIS.AddChampSupValeur ('DATEPIECE', TOBAA.GeTValue('BAA_DATECREAT'));
  TOBDEVIS.AddChampSupValeur ('REFINTERNE', TOBAA.GetString('BAA_INTCODE'));
  TOBDEVIS.AddChampSupValeur ('APPORTEUR', TOBAA.GetString('BAA_APPORTEUR')); 
end;


function TOF_BTIMPORTMULDEV.CreeAppel(TOBAA, TOBA, TOBT: TOB): boolean;
var TOBADR,TOBINTER,TOBADRB :  TOB;
begin
  result := false;
  TOBADR := TOB.Create ('ADRESSES',nil,-1);
  TOBADRB := TOB.Create ('BADRESSES',nil,-1);
  TOBINTER := TOB.Create ('CONTACT',nil,-1);
  //
  TRY
    RechercheOuCreeContact (TOBAA,TOBT,TOBINTER) ;
    TOBA.SetString('AFF_TIERS',TOBT.getString('T_TIERS'));
    TOBA.SetString('AFF_TYPEAFFAIRE','');
    TOBA.SetString('AFF_STATUTAFFAIRE', 'APP');
    TOBA.SetString('AFF_GENERAUTO', 'DIR');
    TOBA.SetString('AFF_PRIOCONTRAT',TOBAA.GetString('BAA_PRIOCONTRAT'));
    TOBA.SetString('AFF_DEVISE',TOBT.GetString('T_DEVISE'));
    TOBA.SetString('AFF_AFFAIREHT',TOBT.GetString('T_FACTUREHT'));
    TOBA.SetString('AFF_REFEXTERNE',TOBAA.GetString('BAA_INTCODE'));
    TOBA.SetString('AFF_LIBELLE',TOBAA.GetString('BAA_LIBELLE'));
    TOBA.SetString('AFF_DESCRIPTIF',TOBAA.GetString('BAA_DESCRIPTIF'));
    TOBA.PutValue ('AFF_DATECREATION',TOBAA.GetValue('BAA_DATECREAT'));
//    TOBA.PutValue ('AFF_DATESOUHAIT',TOBAA.GetValue('BAA_DATESOUHAIT'));
    TOBA.PutValue ('AFF_DATESIGNE',TOBAA.GetValue('BAA_DATESOUHAIT'));
    TOBA.putvalue ('AFF_DATEREPONSE',iDate2099);
    TOBA.SetString('AFF_DOMAINE',TOBAA.GetString('BAA_DOMAINE'));
    TOBA.SetString('AFF_VALLIBRE1',TOBAA.GetString('BAA_VALLIBRE1'));
    TOBA.SetString('AFF_VALLIBRE2',TOBAA.GetString('BAA_VALLIBRE2'));
    TOBA.SetString('AFF_VALLIBRE3',TOBAA.GetString('BAA_VALLIBRE3'));
    TOBA.SetString('AFF_DATELIBRE1',TOBAA.GetString('BAA_DATELIBRE1'));
    TOBA.SetString('AFF_DATELIBRE2',TOBAA.GetString('BAA_DATELIBRE2'));
    TOBA.SetString('AFF_DATELIBRE3',TOBAA.GetString('BAA_DATELIBRE3'));
    TOBA.SetString('AFF_LIBREAFF1',TOBAA.GetString('BAA_LIBREAFF1'));
    TOBA.SetString('AFF_LIBREAFF2',TOBAA.GetString('BAA_LIBREAFF2'));
    TOBA.SetString('AFF_LIBREAFF3',TOBAA.GetString('BAA_LIBREAFF3'));
    TOBA.SetString('AFF_LIBREAFF4',TOBAA.GetString('BAA_LIBREAFF4'));
    TOBA.SetString('AFF_LIBREAFF5',TOBAA.GetString('BAA_LIBREAFF5'));
    TOBA.SetString('AFF_LIBREAFF6',TOBAA.GetString('BAA_LIBREAFF6'));
    TOBA.SetString('AFF_LIBREAFF7',TOBAA.GetString('BAA_LIBREAFF7'));
    TOBA.SetString('AFF_LIBREAFF8',TOBAA.GetString('BAA_LIBREAFF8'));
    TOBA.SetString('AFF_LIBREAFF9',TOBAA.GetString('BAA_LIBREAFF9'));
    TOBA.SetString('AFF_LIBREAFFA',TOBAA.GetString('BAA_LIBREAFFA'));
    TOBA.SetString('AFF_CHARLIBRE1',TOBAA.GetString('BAA_CHARLIBRE1'));
    TOBA.SetString('AFF_CHARLIBRE2',TOBAA.GetString('BAA_CHARLIBRE2'));
    TOBA.SetString('AFF_CHARLIBRE3',TOBAA.GetString('BAA_CHARLIBRE3'));
    TOBA.SetString('AFF_BOOLLIBRE1',TOBAA.GetString('BAA_BOOLLIBRE1'));
    TOBA.SetString('AFF_BOOLLIBRE2',TOBAA.GetString('BAA_BOOLLIBRE2'));
    TOBA.SetString('AFF_BOOLLIBRE3',TOBAA.GetString('BAA_BOOLLIBRE3'));
    TOBA.SetString('AFF_DEMANDEREGLE',TOBAA.GetString('BAA_DEMANDEREGLE'));
    TOBA.SetString('AFF_TVAREDUITE',TOBAA.GetString('BAA_TVAREDUITE'));
    TOBA.SetString('AFF_ETATAFFAIRE',TOBAA.GetString('BAA_ETATAFFAIRE'));
    TOBA.SetString('AFF_RESPONSABLE',TOBAA.GetString('BAA_RESPONSABLE'));
    TOBA.SetString('AFF_BATIMENT',TOBAA.GetString('BAA_BATIMENT'));
    TOBA.SetString('AFF_ETAGE',TOBAA.GetString('BAA_ETAGE'));
    TOBA.SetString('AFF_ESCALIER',TOBAA.GetString('BAA_ESCALIER'));
    TOBA.SetString('AFF_PORTE',TOBAA.GetString('BAA_PORTE'));
    TOBA.SetString('AFF_DEMANDEREGLE',TOBAA.GetString('BAA_DEMANDEREGLE'));
    TOBA.SetString('AFF_TVAREDUITE',TOBAA.GetString('BAA_TVAREDUITE'));
    if (TOBA.getString('AFF_ETATAFFAIRE')='ACA') and (TOBA.getString('AFF_RESPONSABLE')<>'') then TOBA.SetString('AFF_ETATAFFAIRE','AFF');
    TOBA.SetAllModifie(true);
    //
    TOBADR.SetString('ADR_TYPEADRESSE','INT');
    TOBADR.SetInteger('ADR_NADRESSE',0);
    TOBADR.SetInteger('ADR_NUMEROADRESSE',FindNumAdresse+1);
    TOBADR.SetString('ADR_CONTACT',TOBAA.getString('BAA_NOMCONTACT'));
    TOBADR.SetInteger('ADR_NUMEROCONTACT',TOBINTER.getInteger('C_NUMEROCONTACT'));
    if (TOBADR.getString('ADR_CONTACT')='') and (TOBADR.GetInteger('ADR_NUMEROCONTACT')=0) then
    begin
      TOBADR.SetString('ADR_CONTACT','.');
    end;
    TOBADR.SetString('ADR_REFCODE',TOBA.getString('AFF_AFFAIRE'));
    TOBADR.SetString('ADR_JURIDIQUE',TOBAA.getString('BAA_JURIDIQUE'));
    TOBADR.SetString('ADR_LIBELLE',TOBAA.getString('BAA_LIBADR'));
    TOBADR.SetString('ADR_LIBELLE2',TOBAA.getString('BAA_LIBADR2'));
    TOBADR.SetString('ADR_ADRESSE1',TOBAA.getString('BAA_ADRESSE1'));
    TOBADR.SetString('ADR_ADRESSE2',TOBAA.getString('BAA_ADRESSE2'));
    TOBADR.SetString('ADR_ADRESSE3',TOBAA.getString('BAA_ADRESSE3'));
    TOBADR.SetString('ADR_CODEPOSTAL',TOBAA.getString('BAA_CODEPOSTAL'));
    TOBADR.SetString('ADR_VILLE',TOBAA.getString('BAA_VILLE'));
    TOBADR.SetString('ADR_TELEPHONE',TOBAA.getString('BAA_TELEPHONE'));
    TOBADR.SetString('ADR_FAX',TOBAA.getString('BAA_MOBILE'));

    TOBADR.putValue('ADR_BLOCNOTE',TOBAA.getString('BAA_BLOCNOTE'));
    TOBADR.SetAllModifie(true);

    TOBADRB.SetString('BA0_TYPEADRESSE','INT');
    TOBADRB.SetInteger('BA0_NUMEROADRESSE',TOBADR.GetInteger('ADR_NUMEROADRESSE'));
    TOBADRB.SetString('BA0_GARDIEN',TOBAA.getString('BAA_GARDIEN'));
    TOBADRB.SetString('BA0_CODEENTREE1',TOBAA.getString('BAA_CODEENTREE1'));
    TOBADRB.SetString('BA0_CODEENTREE2',TOBAA.getString('BAA_CODEENTREE2'));
    TOBADRB.SetString('BA0_CODEENTREE2',TOBAA.getString('BAA_CODEENTREE2'));
    TOBADRB.SetString('BA0_INT','X');


    result := EcritContact (TOBINTER,TOBT);
    if not result then
    begin
      AddReport ('NOK','Création/mise à jour contact',TOBAA);
      exit;
    end;
    result := TOBADR.InsertDB(nil);
    if not result then
    begin
      AddReport ('NOK','Création de l''adresse d''intervention',TOBAA);
      exit;
    end;
    result := TOBADRB.InsertDB(nil);
    if not result then
    begin
      AddReport ('NOK','Création de l''adresse d''intervention BTP',TOBAA);
      exit;
    end;
    result := TOBA.InsertDB(nil);
    if not result then
    begin
      AddReport ('NOK','Création de l''intervention',TOBAA);
      exit;
    end;
    //
  FINALLY
    TOBADR.free;
    TOBADRB.free;
    TOBINTER.free;
  END;
end;


function  TOF_BTIMPORTMULDEV.FindNumAdresse : integer;
var QQ : Tquery;
begin
  result := 0;
  QQ := OpenSql ('SELECT MAX(ADR_NUMEROADRESSE) FROM ADRESSES', True);
  if not QQ.eof then
  begin
    result := QQ.fields[0].AsInteger;
  end;
  ferme (QQ);
end;

function  TOF_BTIMPORTMULDEV.FindNumContact (TOBT : TOB) : integer;
var QQ : Tquery;
begin
  result := 0;
  QQ := OpenSql ('SELECT MAX(C_NUMEROCONTACT) FROM CONTACT WHERE C_TYPECONTACT="T" AND C_AUXILIAIRE = "'+TOBT.GetString('T_AUXILIAIRE')+'"', True);
  if not QQ.eof then
  begin
    result := QQ.fields[0].AsInteger;
  end;
  ferme (QQ);
end;


procedure TOF_BTIMPORTMULDEV.RechercheOuCreeContact(TOBAA, TOBT, TOBINTER: TOB);
var Sql : String;
    QQ : TQuery;
    Okok : boolean;
begin
  if TOBAA.getString('BAA_NOMCONTACT')='' then exit;
  okok := false;
  Sql := 'SELECT *,"X" AS EXISTANT  FROM CONTACT '+
         'WHERE '+
         'C_TYPECONTACT="T" AND '+
         'C_AUXILIAIRE="'+TOBT.GetString('T_AUXILIAIRE')+'" AND '+
         'C_NOM="'+TOBAA.getString('BAA_NOMCONTACT')+'" AND '+
         'C_PRENOM="'+TOBAA.getString('BAA_PRENOMCONTACT')+'"';
  QQ := OpenSql (Sql,true,1,'',true);
  if not QQ.eof then
  begin
    Okok := true;
    TOBINTER.SelectDB('',QQ);
    // pour le cas ou
    TOBINTER.SetString('C_CIVILITE',TOBAA.GetString('BAA_CIVCONTACT'));
    TOBINTER.SetString('C_TELEPHONE',TOBAA.GetString('BAA_TELCONTACT'));
    TOBINTER.SetString('C_CLETELEPHONE',CleTelephone(TOBAA.GetString('BAA_TELCONTACT')));
    TOBINTER.SetString('C_TELEX',TOBAA.GetString('BAA_MOBILE'));
    TOBINTER.SetString('C_CLETELEX',CleTelephone(TOBAA.GetString('BAA_MOBILE')));
    TOBINTER.SetString('C_RVA',TOBAA.GetString('BAA_MAILCONTACT'));
  end;
  ferme (QQ);
  if not okok then
  begin
    TOBINTER.AddChampSupValeur ('EXISTANT','-'); 
    TOBINTER.SetString('C_TYPECONTACT','T');
    TOBINTER.SetString('C_AUXILIAIRE',TOBT.GetString('T_AUXILIAIRE'));
    TOBINTER.SetString('C_TIERS',TOBT.GetString('T_TIERS'));
    TOBINTER.SetString('C_NOM',TOBAA.GetString('BAA_NOMCONTACT'));
    TOBINTER.SetString('C_PRENOM',TOBAA.GetString('BAA_PRENOMCONTACT'));
    TOBINTER.SetString('C_NATUREAUXI',TOBT.GetString('T_NATUREAUXI'));
    TOBINTER.SetString('C_CIVILITE',TOBAA.GetString('BAA_CIVCONTACT'));
    TOBINTER.SetString('C_TELEPHONE',TOBAA.GetString('BAA_TELCONTACT'));
    TOBINTER.SetString('C_CLETELEPHONE',CleTelephone(TOBAA.GetString('BAA_TELCONTACT')));
    TOBINTER.SetString('C_TELEX',TOBAA.GetString('BAA_MOBILE'));
    TOBINTER.SetString('C_CLETELEX',CleTelephone(TOBAA.GetString('BAA_MOBILE')));
    TOBINTER.SetString('C_RVA',TOBAA.GetString('BAA_MAILCONTACT'));
    TOBINTER.SetString('C_SOCIETE',V_PGI.CodeSociete);
    TOBINTER.SetInteger('C_NUMEROCONTACT',FindNumContact(TOBT)+1);
  end;
end;

function TOF_BTIMPORTMULDEV.EcritContact(TOBINTER,TOBT: TOB): boolean;
var II : integer;
begin
  result := true;
  if TOBINTER.GetString('C_NOM')='' then exit;
  II := 1;
  if TOBINTER.GetString('EXISTANT')='X' then
  begin
    TOBINter.UpdateDB(false); 
  end else
  begin
    repeat
      TRY
        TOBINTER.SetAllModifie(true);
        TOBINter.InsertDB(nil);
        break;
      EXCEPT
        TOBINTER.SetInteger('C_NUMEROCONTACT',FindNumContact(TOBT)+1);
        inc (II);
      end;
    until II > 100;
    if II> 100 then result := false;
  end;
end;

function TOF_BTIMPORTMULDEV.PositionneTraite(TOBAA: TOB): boolean;
begin
  result := (ExecuteSQL('UPDATE BAFFAIREATT SET BAA_TRAITE="X" WHERE BAA_GUID="'+TOBAA.getString('BAA_GUID')+'"') > 0);
  if not result then  AddReport ('NOK','Mise à jour de l''enregistrement '+TOBAA.getString('BAA_GUID')+' impossible',TOBAA);
end;

function TOF_BTIMPORTMULDEV.InsereDatas(TOBA, TOBAA, TOBT: TOB): boolean;

  function constitueDataIns (TOBA,TOBL,TOBID : TOB) : boolean;
  var TOBI : TOB;
      TypeData : string;
  begin
    result := true;
    TypeData :=TOBL.getString('BIA_QUALIFIANTBLOB');
    if Pos(TypeData,'IMG;IMA;SIA;SIG') = 0 then BEGIN result:= false; exit; END;
    TOBI := TOB.Create ('LIENSOLE',TOBID,-1);
    if TypeData = 'IMG' then
    begin
      TOBI.setString('LO_TABLEBLOB','APP');
      TOBI.setString('LO_QUALIFIANTBLOB','PHJ');
      TOBI.setString('LO_EMPLOIBLOB','IMG');
    end else if TypeData = 'IMA' then
    begin
      TOBI.setString('LO_TABLEBLOB','APL');
      TOBI.setString('LO_QUALIFIANTBLOB','PHJ');
      TOBI.setString('LO_EMPLOIBLOB','IMG');
    end else if TypeData = 'SIA' then
    begin
      TOBI.setString('LO_TABLEBLOB','APP');
      TOBI.setString('LO_QUALIFIANTBLOB','COM');
      TOBI.setString('LO_EMPLOIBLOB','REA');
    end else if TypeData = 'SIG' then
    begin
      TOBI.setString('LO_TABLEBLOB','APP');
      TOBI.setString('LO_QUALIFIANTBLOB','PHJ');
      TOBI.setString('LO_EMPLOIBLOB','SIG');
    end;
    //
    TOBI.SetString('LO_LIBELLE',TOBL.getString('BIA_LIBELLE'));
    TOBI.PutValue('LO_OBJET',TOBL.Getvalue('BIA_OBJET'));
    TOBI.SetInteger('LO_RANGBLOB',TOBL.GetInteger('BIA_RANGBLOB'));
    TOBI.SetString('LO_IDENTIFIANT',TOBA.GetString('AFF_AFFAIRE'));
  end;

var TOBDATAS,TOBID : TOB;
    QQ : TQuery;
    II : integer;
    TOBL : TOB;
begin
  result := true;
  TOBDATAS := TOB.create ('LES DATAS',nil,-1);
  TOBID := TOB.create ('LES DATAS INT',nil,-1);
  TRY
    QQ := OpenSql ('SELECT * FROM BDATAATT WHERE BIA_GUID="'+TOBAA.getString('BAA_GUID')+'"',true,-1,'',true);
    if not QQ.eof then
    begin
      TOBDATAS.LoadDetailDB('BDATAATT','','',QQ,false);
      for II := 0 to TOBDATAS.detail.count -1 do
      begin
        TOBL := TOBDATAS.detail[II];
        if not constitueDataIns (TOBA,TOBL,TOBID) then
        begin
          AddReport ('NOK','Ajout de DATAS '+TOBL.GetString('BIA_QUALIFIANTBLOB')+':'+TOBL.GetString('BIA_RANGBLOB')+' pour l''enregistrement '+TOBAA.getString('BAA_GUID')+' impossible',TOBAA);
          result := false;
          break;
        end;
      end;
      if result then
      begin
        TOBID.SetAllModifie(true); 
        if not TOBID.InsertDB(nil) then
        begin
          AddReport ('NOK','Ajout de DATAS : insertion impossible',TOBAA);
          result := false;
        end;
      end;
    end;
  FINALLY
    TOBDATAS.free;
    TOBID.free;
  END;
end;

Initialization
  registerclasses ( [ TOF_BTIMPORTMULDEV ] ) ;
end.
