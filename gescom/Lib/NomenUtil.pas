unit NomenUtil;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
    UTOB, Hctrls, SysUtils, Math,
{$IFDEF EAGLCLIENT}
{$ELSE}
    {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
{$IFDEF BTP}
    UtilTarif,
{$ENDIF}
    SaisUtil, comctrls, HEnt1, UtilArticle, HDimension,Factcalc,UtilPgi,ENtGC,UentCommun;

//  Ordre des valorisations d'une nomenclature dans le tableau des valeurs
Type T_TypeVal = (valDPA, valDPR, valPVHT, valPVTTC, valPAHT, valPRHT, valPMAP, valPMRP, valSup,ValTps) ;
//
// Modif BTP
//
Type T_ExceptVal = (ExPvHT,ExPvUniq,ExAucun);
// -----
//  Fonctions liées aux nomenclatures
procedure ChargeNomen(CodeNomen : string; Complet : Boolean; var TobNomen : TOB);
function  Circularite(CodeNomen : string; TobNomen : TOB; var TOBErreur : TOB) : Boolean;
procedure MiseAPlat(CodeNomen : string; TobNomen : TOB; var TobResul : TOB; Complet : Boolean = False );
procedure APlat(TobNomen : TOB; var TobResul : TOB; Complet : Boolean = False );
function  Valorise(CodeNomen : string; TobNomen : TOB; var Resul : T_Valeurs) : Boolean;
//  Fonctions liées aux cas d'emploi
procedure CasEmploi(CodeArticle : string; Colonne : string; Complet : Boolean; var TobResul : TOB);
//  Fonctions liées aux Treeview
procedure AffecteLibelle (var TobNomen : TOB; Colonne : string);
procedure AfficheTreeView(TV : TTreeView; TOBTemp : TOB; Colonne : string);
(*function  ChercheNode(TV : TTreeView; TN : TTreeNode; Valeur : string; // DBR Fiche 10995
                      Niveau : integer = 1) : TTreeNode; *)
function  ChercheNode(TV : TTreeView; TN : TTreeNode; Valeur : integer; // DBR Fiche 10995
                      Niveau : integer = 1) : TTreeNode;
function  FindParentNiv1(TN : TTreeNode) : TTreeNode;
//  Divers
procedure LibelleDimensions(CodeArticle : string; TOBArt : TOB; var LibelleDim, GrilleDim : string; Abrege : Boolean = False);
//
// Modif BTP
//
procedure ValoriseOuvrage (CodeNomen : string;TOBMereEnt,TOBMereLig : TOB; var Resul : T_Valeurs;AvecPourcent:boolean=true;RecupEntete:boolean=false);
procedure renseigneValoArt (TOBIndice,TOBART,TOBMereLig : TOB; SaisieComposant : Boolean=true);
procedure CalculeOuvrage (CodeNomen : String; TObMereEntete,TOBMereLigne : TOB; Exception: T_ExceptVal ;var Resul : T_Valeurs;AvecPourcent:Boolean=true);
procedure ChargeEntete (CodeNomen : String; var TOBEnt : Tob);
procedure CumuleElementTable (TOBDetail : TOB ; Exception : T_ExceptVal; var resul:T_Valeurs);
Function UniteOuvrage (article : string): String;
procedure FormatageTableau (var resul : T_valeurs;Decimale:integer=-1);
procedure InitTableau (var resul : T_valeurs);
function CalculSurTableau ( operation : String; entree : T_valeurs; operande : T_valeurs): T_valeurs; overload
function CalculSurTableau ( operation : String; entree : T_valeurs; operande : double): T_valeurs; overload
procedure RepriseDonneeArticle (TOBLN,TOBArticle : TOB; FromCopierColler : boolean = false; WithValorisation : boolean= true);
procedure ReajusteLigneParDoc (TOBLD,TOBLN: TOB;DEV:RDevise;Niveau : integer);
{$IFDEF BTP}
procedure RecupValorisation (TOBREF,TOBLIG,TOBMereLig : TOB);
procedure RecupValoHrs (TOBRef,TOBMereLig : TOB);
function ExistNomen (Article,Nomen : string; TOBREF : TOB = nil) : boolean;
function AffecteCodeNomen (CodeArticle,CodeNomenRef : string; TOBREF : TOB = nil) : string;
function CompteDeclinaison(Article: string): integer;
{$ENDIF}
procedure AppliqueValoDomaineAct (TOBIndice,TOBART : TOB);

//procedure ChampSupNomen ( TOBNomen : TOB;); // DBR Fiche 10995
procedure ChampSupNomen ( TOBNomen : TOB; ToutesLesFillesLibelle : boolean = True); // DBR Fiche 10995

implementation
uses factvariante
,factouvrage
,Paramsoc
;
//procedure ChampSupNomen ( TOBNomen : TOB); // DBR Fiche 10995
procedure ChampSupNomen ( TOBNomen : TOB; ToutesLesFillesLibelle : boolean = True); // DBR Fiche 10995
begin
if (not TOBNomen.FieldExists ('BNP_TYPERESSOURCE')) then TOBNomen.AddChampSupValeur ('BNP_TYPERESSOURCE','',false);
//
if (VarIsNUll(TOBNomen.GetValue ('BNP_TYPERESSOURCE'))) or (VarAsType(TOBNomen.GetValue ('BNP_TYPERESSOURCE'),VarString)=#0) then
begin
	TOBNomen.Putvalue ('BNP_TYPERESSOURCE','');
end;
//
if (not TOBNomen.FieldExists ('GNE_DOMAINE'))       then TOBNomen.AddChampSupValeur ('GNE_DOMAINE','',false);
if (not TOBNomen.FieldExists ('GA_TYPEARTICLE'))    then TOBNomen.AddChampSupValeur ('GA_TYPEARTICLE','',false);
if (not TOBNomen.FieldExists ('GA_FOURNPRINC'))     then TOBNomen.AddChampSupValeur ('GA_FOURNPRINC','',false);
//FV1 : 18/02/2015 - Gestion du Prix d'achat dans la fiche article
if (not TOBNomen.FieldExists ('GA_PAUA'))           then TOBNomen.AddChampSupValeur ('GA_PAUA',0,false);
if (not TOBNomen.FieldExists ('GA_QUALIFUNITEACH')) then TOBNomen.AddChampSupValeur ('GA_QUALIFUNITEACH',0,false);
//
if (not TOBNomen.FieldExists ('GA_PAHT'))           then TOBNomen.AddChampSupValeur ('GA_PAHT',0,false);
if (not TOBNomen.FieldExists ('GA_PVHT'))           then TOBNomen.AddChampSupValeur ('GA_PVHT',0,false);
if (not TOBNomen.FieldExists ('GA_NATUREPRES'))     then TOBNomen.AddChampSupValeur ('GA_NATUREPRES','',false);
if (not TOBNomen.FieldExists ('GA2_SOUSFAMTARART')) then TOBNomen.AddChampSupValeur ('GA2_SOUSFAMTARART','',false);
//
if (VarIsNUll(TOBNomen.GetValue ('GA2_SOUSFAMTARART'))) or (VarAsType(TOBNomen.GetValue ('GA2_SOUSFAMTARART'),VarString)=#0) then
begin
	TOBNomen.Putvalue ('GA2_SOUSFAMTARART','');
end;
// -------- Champs Sup Pur -----------------
if (not TOBNomen.FieldExists ('GA_HEURE')) then TOBNomen.AddChampSupValeur ('GA_HEURE',0,false);
if (not TOBNomen.FieldExists ('GNL_PRIXPOURQTE')) then TOBNomen.AddChampSupValeur ('GNL_PRIXPOURQTE',1,false);
//  Ajout du champ sup. LIBELLE = CODEARTICLE (QUANTITE)
if (not TOBNomen.FieldExists ('LIBELLE')) then TobNomen.AddChampSupValeur('LIBELLE','', ToutesLesFillesLibelle); // DBR Fiche 10995
{$IFDEF BTP}
if (not TOBNomen.FieldExists ('QTEDUDETAIL')) then TobNomen.AddChampSupValeur('QTEDUDETAIL', 1,false);
if (not TOBNomen.FieldExists ('UNITENOMENC')) then TobNomen.AddChampSupValeur('UNITENOMENC', '',false);
if (not TOBNomen.FieldExists ('DPA')) then TOBNomen.AddChampSupValeur ('DPA',0,false);
if (not TOBNomen.FieldExists ('COEFFG')) then TOBNomen.AddChampSupValeur ('COEFFG',0,false);
if (not TOBNomen.FieldExists ('DPR')) then TOBNomen.AddChampSupValeur ('DPR',0,false);
if (not TOBNomen.FieldExists ('COEFMARG')) then TOBNomen.AddChampSupValeur ('COEFMARG',0,false);
if (not TOBNomen.FieldExists ('PVHT')) then TOBNomen.AddChampSupValeur ('PVHT',0,false);
if (not TOBNomen.FieldExists ('PVTTC')) then TOBNomen.AddChampSupValeur ('PVTTC',0,false);
if (not TOBNomen.FieldExists ('PMAP')) then TOBNomen.AddChampSupValeur ('PMAP',0,false);
if (not TOBNomen.FieldExists ('PMRP')) then TOBNomen.AddChampSupValeur ('PMRP',0,false);
if (not TOBNomen.FieldExists ('TPSUNITAIRE')) then TOBNomen.AddChampSupValeur ('TPSUNITAIRE',0,false);
if (not TOBNomen.FieldExists ('GCA_PRIXBASE')) then TOBNomen.AddChampSupValeur ('GCA_PRIXBASE',0,false);
if (not TOBNomen.FieldExists ('GCA_QUALIFUNITEACH')) then TOBNomen.AddChampSupValeur ('GCA_QUALIFUNITEACH','',false);
if (not TOBNomen.FieldExists ('GF_CALCULREMISE')) then TOBNomen.AddChampSupValeur ('GF_CALCULREMISE','',false);
if (not TOBNomen.FieldExists ('GF_PRIXUNITAIRE')) then TOBNomen.AddChampSupValeur ('GF_PRIXUNITAIRE',0,false);
if (not TOBNomen.FieldExists ('TGA_FOURNPRINC')) then TOBNomen.AddChampSupValeur ('TGA_FOURNPRINC','',false);
if (not TOBNomen.FieldExists ('TGA_PRESTATION')) then TOBNomen.AddChampSupValeur ('TGA_PRESTATION','',false);
// --
if (not TOBNomen.FieldExists ('GCA_PRIXVENTE')) then TOBNomen.AddChampSupValeur ('GCA_PRIXVENTE',0,false);

{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : B. Bicheray

Créé le ...... : 28/01/2000

Modifié le ... :   /  /

Description .. : Chargement d'une nomenclature en TOB

Mots clefs ... : NOMENCLATURE

*****************************************************************}

//=============================================================================
//  Chargement d'une nomenclature  (BBY - 22/12/99)
//  Parametres : CodeNomen : code de la nomenclature a charger. obligatoire.
//               Complet   : Vrai indique le chargement recursif des sous niveaux
//               TobNomen  : Tob dans laquelle doit se faire le chargement. elle
//                           est cree si necessaire.
//=============================================================================
procedure ChargeNomen(CodeNomen : string; Complet : Boolean; var TobNomen : TOB);
var
		Valeurs : T_Valeurs;
    i_ind1,Indice : integer;
    Select : string;
    Q : TQuery;
    TobFille : TOB;
{$IFDEF BTP}
    TobEnt: TOB;
    QEnt : Tquery;
{$ENDIF}
begin
  //  CodeNomen obligatoire
  if CodeNomen = '' then Exit;
  //  TobNomen non renseignée, on la crée
  if TobNomen = nil then TobNomen := TOB.Create('', nil, -1);
  //  Chargement du niveau courant de nomenclature
//
  Select := 'SELECT NOMENLIG.*,N.BNP_TYPERESSOURCE,GNE_DOMAINE,GNE_QTEDUDETAIL AS QTEDUDETAIL,'
  					+	'GA_TYPEARTICLE,GA_FOURNPRINC,GA_PAUA,GA_PAHT,GA_PVHT,GA_QUALIFUNITEACH, GA_NATUREPRES,'
            + 'GA2_SOUSFAMTARART FROM NOMENLIG '
            + 'LEFT JOIN NOMENENT ON GNE_NOMENCLATURE=GNL_NOMENCLATURE '
            + 'LEFT JOIN ARTICLE ON GA_ARTICLE=GNL_ARTICLE '
            + 'LEFT JOIN ARTICLECOMPL ON GA2_ARTICLE=GNL_ARTICLE '
            + 'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=GA_NATUREPRES '
            + 'WHERE GNL_NOMENCLATURE = "' + CodeNomen + '" ORDER BY GNL_NUMLIGNE';
//
(*
	Select := 'Select NOMENLIG.*,GNE_DOMAINE from NOMENLIG '+
            'LEFT JOIN NOMENENT ON GNE_NOMENCLATURE=GNL_NOMENCLATURE '+
            'where GNL_NOMENCLATURE = "' + CodeNomen + '" ORDER BY GNL_NUMLIGNE';
*)
  Q := OpenSQL(Select, True,-1, '', True);
  if not Q.EOF then
  begin
  	TobNomen.LoadDetailDB('NOMENLIG', '', '', Q, False);
  end;
  Ferme(Q);

  //  Si Complet a vrai, on balaye les filles pour charger les eventuelles nomenclatures
  //  de sous niveau
  if Complet then
  begin
    for i_ind1 := 0 to TobNomen.Detail.Count - 1 do
    begin
      TobFille := TobNomen.Detail[i_ind1];
      ChampSupNomen (TOBFille);
      {$IFDEF BTP}
      Tobfille.putvalue('QTEDUDETAIL',1);
      Tobfille.putvalue('UNITENOMENC','');
      if TOBNomen.fieldExists('GNE_DOMAINE') then
      begin
        Tobfille.putvalue('GNE_DOMAINE',TOBNomen.getValue('GNE_DOMAINE'));
      end;

      if TOBNomen.fieldExists('GA_FOURNPRINC') then
      begin
        Tobfille.putvalue('GA_FOURNPRINC',TOBNomen.getValue('GA_FOURNPRINC'));
      end;

      if TobFille.getvalue('GNL_SOUSNOMEN') <> '' then
      begin
        TOBEnt:= TOB.create('NOMENENT',nil,-1);
        Select := 'SELECT GNE_QTEDUDETAIL FROM NOMENENT WHERE GNE_NOMENCLATURE ="' + Tobfille.getvalue('GNL_SOUSNOMEN') + '"';
        QEnt:= Opensql (Select,true,-1, '', True);
        if not QEnt.eof then
        begin
          TobEnt.selectdb ('',QEnt,true);
          Tobfille.putvalue('QTEDUDETAIL',TOBEnt.getvalue('GNE_QTEDUDETAIL'));
          TOBEnt.Free;
        end;
        ferme(QEnt);
      end ;
      {$ENDIF}
      ChargeNomen(TobFille.GetValue('GNL_SOUSNOMEN'), Complet, TobFille);
      if TOBFille.detail.count > 0 then
      begin
      	InitTableau (Valeurs);
        for Indice := 0 to TOBFille.detail.count -1 do
        begin
          CumuleElementTable (TOBFille.detail[Indice],ExAucun,valeurs);
        end;
        TOBFille.putvalue ('DPA',         valeurs[0]);
        TOBFille.putvalue ('DPR',         valeurs[1]);
        TOBFille.putvalue ('PVHT',        valeurs[2]);
        TOBFille.putvalue ('PVTTC',       valeurs[3]);
        TOBFille.putvalue ('PMAP',        valeurs[6]);
        TOBFille.putvalue ('PMRP',        valeurs[7]);
        TOBFille.putvalue ('TPSUNITAIRE', valeurs[9]);
        //
        TOBFille.putValue('GA_PAHT',         TOBFille.getValue('DPA'));
        TOBFille.putValue('GA_PVHT',         TOBFille.getValue('PVHT'));
        TOBFille.putvalue ('GF_PRIXUNITAIRE',TOBFille.GetValue('PVHT'));
      end;
    end;
  end else
  begin
    for i_ind1 := 0 to TobNomen.Detail.Count - 1 do
    begin
      TobFille := TobNomen.Detail[i_ind1];
      ChampSupNomen (TOBFille);
    end;
  end;
  //
  if TOBNomen.detail.count > 0 then
  begin
    InitTableau (Valeurs);
    for Indice := 0 to TOBNomen.detail.count -1 do
    begin
      //if TOBNomen.fieldExists('QTEDUDETAIL') then
      //  TOBNomen.detail[Indice].putvalue('QTEDUDETAIL',TOBNomen.getValue('QTEDUDETAIL'));
      CumuleElementTable (TOBNomen.detail[Indice],ExAucun,valeurs);
    end;
    TOBNomen.putvalue ('DPA',valeurs[0]);
    TOBNomen.putvalue ('DPR',valeurs[1]);
    TOBNomen.putvalue ('PVHT',valeurs[2]);
    TOBNomen.putvalue ('PVTTC',valeurs[3]);
    TOBNomen.putvalue ('PMAP',valeurs[6]);
    TOBNomen.putvalue ('PMRP',valeurs[7]);
    TOBNomen.putvalue ('TPSUNITAIRE',valeurs[9]);
    //
    TOBNomen.putValue('GA_PAHT',TOBNomen.getValue('DPA'));
    TOBNomen.putValue('GA_PVHT',TOBNomen.getValue('PVHT'));
    TOBNomen.putvalue ('GF_PRIXUNITAIRE',TOBNomen.GetValue('PVHT'));
  end;
  //
end;

{***********A.G.L.***********************************************
Auteur  ...... : B. Bicheray

Créé le ...... : 28/01/2000

Modifié le ... :   /  /

Description .. : Vérification de non circularité des appels d'une nomenclature

Mots clefs ... : NOMENCLATURE, CIRCULARITE

*****************************************************************}

//=============================================================================
//  Verification de non circularité  (BBY - 29/12/99)
//  Parametres : CodeNomen : code de la nomenclature a verifier. obligatoire.
//               TobNomen  : Tob dans laquelle se trouve la nomenclature.
//=============================================================================
function Circularite(CodeNomen : string; TobNomen : TOB; var TOBErreur : TOB) : Boolean;
var
    i_ind1 : integer;
    TOBTemp, TOBTempErr : TOB;
begin
Result := True;
for i_ind1 := 0 to TobNomen.Detail.Count - 1 do
    begin
    TOBTemp := TobNomen.Detail[i_ind1];
//  pas de code de sous nomenclature, on continue
    if TOBTemp.GetValue('GNL_SOUSNOMEN') = '' then Continue;
//  on compare avec le code de sous nomenclature eventuel
    if trim(CodeNomen) = trim(TOBTemp.GetValue('GNL_SOUSNOMEN')) then
        begin
        TOBTempErr := TOB.Create('', TOBErreur, -1);
        TOBTempErr.AddChampSup('GNL_SOUSNOMEN', True);
        TOBTempErr.PutValue('GNL_SOUSNOMEN', CodeNomen);
        Result := False;
//        Exit;
        end
    else
        begin
        TOBTempErr := TOB.Create('', TOBErreur, -1);
        TOBTempErr.AddChampSup('GNL_SOUSNOMEN', True);
        TOBTempErr.PutValue('GNL_SOUSNOMEN', TOBTemp.GetValue('GNL_SOUSNOMEN'));
        if not Circularite(CodeNomen, TOBTemp, TOBTempErr) then
            begin
            Result := False;
//            Exit;
            end;
        end;
    end;
end;

//=============================================================================
//  Mise à plat d'une nomenclature  (BBY - 04/01/2000)
//  Parametres : CodeNomen : code de la nomenclature à mettre à plat. obligatoire.
//               TobNomen  : Tob contenant la nomenclature. si elle est vide on
//                           fait par défaut un chargement complet à partir du
//                           parametre CodeNomen
//               TobResul  : Tob résultat de la mise à plat.
//                           Contient les colonnes de la table NOMENLIG plus LIBELLE
//               Complet   : Vrai indique la mise à plat détaillée avec conservation
//                           de toutes les références composants. Faux indique la
//                           mise à plat avec une ligne par composant.
//=============================================================================
{***********A.G.L.***********************************************
Auteur  ...... : B. Bicheray

Créé le ...... : 28/01/2000

Modifié le ... :   /  /    

Description .. : Chargement en tob  de tous les composants sur un seul niveau

Mots clefs ... : NOMENCLATURE, MISE À PLAT, COMPOSANTS

*****************************************************************}

procedure MiseAPlat(CodeNomen : string; TobNomen : TOB; var TobResul : TOB; Complet : Boolean = False );
var
    i_ind1 : integer;
    TobTemp : TOB;
begin
//  CodeNomen obligatoire
if CodeNomen = '' then Exit;
//  TobNomen non renseignée, on la charge
if TobNomen = nil then ChargeNomen(CodeNomen, True, TobNomen);
//  Remise à blanc de la Tob resultat
if TobResul = nil then
    TobResul := TOB.Create('', nil, -1)
else
    TobResul.ClearDetail;
//  Mise à plat
for i_ind1 := 0 to TobNomen.Detail.Count - 1 do
    begin
    TobTemp := TobNomen.Detail[i_ind1];
    APlat(TobTemp, TobResul, Complet);
    end;
end;

procedure APlat(TobNomen : TOB; var TobResul : TOB; Complet : Boolean = False );
var
    i_ind1 : integer;
    TobTemp : TOB;
begin
if TobNomen.Detail.Count <> 0 then
//  il y a un detail donc c'est une nomenclature --> appel recursif sur chaque fille
    begin
    for i_ind1 := 0 to TobNomen.Detail.Count - 1 do
        begin
        TobTemp := TobNomen.Detail[i_ind1];
        APlat(TobTemp, TobResul, Complet);
        end;
    end
    else
//  Pas de détail donc on insere dans la tob
    begin
//  Complet a True, on ajoute systematiquement une ligne
    if Complet then
        begin
        TobTemp := TOB.Create('NOMENLIG', TobResul, -1);
        TobTemp.Dupliquer(TobNomen, False, True);
        end
        else
//  Complet a False, on Recherche le composant dans la tob resultat et on ajoute la qte
        begin
        TobTemp := TobResul.FindFirst(['GNL_ARTICLE'], [TobNomen.GetValue('GNL_ARTICLE')], False);
        if TobTemp = nil then
            begin
            TobTemp := TOB.Create('NOMENLIG', TobResul, -1);
            TobTemp.Dupliquer(TobNomen, False, True);
            end
            else
            begin
            TobTemp.PutValue('GNL_QTE', TobTemp.GetValue('GNL_QTE') + TobNomen.GetValue('GNL_QTE'));
            end;
        end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : B. Bicheray

Créé le ...... : 31/01/2000

Modifié le ... :   /  /    

Description .. : Calcul de la valeur d'une nomenclature en fonction des divers prix de ses composants

Mots clefs ... : NOMENCLATURE, PRIX, COMPOSANT

*****************************************************************}

//=============================================================================
//  Valorisation d'une nomenclature  (BBY - 31/01/2000)
//  Parametres : CodeNomen  : code de la nomenclature à valoriser. obligatoire si la tob est vide.
//                            Prioritaire sur la tob si les deux sont renseignés
//               TobNomen   : Tob contenant la nomenclature. obligatoire si le code est vide.
//               Resul      : tableau contenant les differentes valorisations
//                            poste 0 : Dernier prix d'achat
//                            poste 1 : Dernier prix de revient
//                            poste 2 : Prix de vente HT
//                            poste 3 : Prix de vente TTC
//                            poste 4 : Prix d'achat HT
//                            poste 5 : Prix de revient HT
//                            poste 6 : Prix moyen d'achat pondéré
//                            poste 7 : Prix moyen de revient pondéré
//=============================================================================
function Valorise(CodeNomen : string; TobNomen : TOB; var Resul : T_Valeurs) : boolean;
var
    i_ind1 : integer;
    d_temp1 : double;
    TobArt : TOB;
    TobResul : TOB;
begin
//  on a au moins le code ou la tob renseignés
Result := True;
if (CodeNomen = '') and (TobNomen = nil) then
    begin
    Result := False;
    Exit;
    end;
//  init du tableaux résultat
for i_ind1 := 0 to 9 do Resul[i_ind1] := 0.0;
//  le code nomenclature est prioritaire sur la tob si les deux existent.
if CodeNomen <> '' then
    begin
    ChargeNomen(CodeNomen, True, TobNomen);
//  mise à plat de la nomenclature
    TobResul := TOB.Create('', nil, -1);
    MiseAPlat(CodeNomen, TobNomen, TobResul, True);
    end
    else
    begin
    TobResul := TobNomen;
    end;
//  Calcul des valorisation de la nomenclature
TobArt := TOB.Create('ARTICLE', nil, -1);
for i_ind1 := 0 to TobResul.Detail.Count - 1 do
    begin
    d_temp1 := TobResul.Detail[i_ind1].GetValue('GNL_QTE');
    TobArt.PutValue('GA_ARTICLE', TobResul.Detail[i_ind1].GetValue('GNL_ARTICLE'));
    TobArt.ChargeCle1;
    TobArt.LoadDB;
    Resul[0] := Resul[0] + (TobArt.GetValue('GA_DPA')   * d_temp1);
    Resul[1] := Resul[1] + (TobArt.GetValue('GA_DPR')   * d_temp1);
    Resul[2] := Resul[2] + (TobArt.GetValue('GA_PVHT')  * d_temp1);
    Resul[3] := Resul[3] + (TobArt.GetValue('GA_PVTTC') * d_temp1);
    Resul[4] := Resul[4] + (TobArt.GetValue('GA_PAHT')  * d_temp1);
    Resul[5] := Resul[5] + (TobArt.GetValue('GA_PRHT')  * d_temp1);
    Resul[6] := Resul[6] + (TobArt.GetValue('GA_PMAP')  * d_temp1);
    Resul[7] := Resul[7] + (TobArt.GetValue('GA_PMRP')  * d_temp1);
    end;
if CodeNomen <> '' then
    begin
    TobResul.Free;
    TobResul := nil;
    end;
TobArt.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : B. Bicheray

Créé le ...... : 28/01/2000

Modifié le ... :   /  /

Description .. : recherche cas d'emploi d'un article

Mots clefs ... : NOMENCLATURE, CAS D'EMPLOI, COMPOSANT

*****************************************************************}

//=============================================================================
//  Recherche des cas d'emploi  (BBY - 04/01/2000)
//  Parametres : CodeArticle : code de l'article dont on cherche les cas d'emploi. obligatoire.
//               Colonne     : permet de preciser une colonne particuliere de recherche
//                             (CODEARTICLE, SOUSNOMEN, NOMENCLATURE). par defaut c'est toutes
//               Complet     : precise si on veut toute l'arborescence d'emploi de l'article
//                             ou seulement le premier niveau
//               TobResul    : Tob résultat de la recherche.
//                             Contient les colonnes de la table NOMENLIG plus LIBELLE
//=============================================================================
procedure CasEmploi(CodeArticle : string; Colonne : string; Complet : Boolean; var TobResul : TOB);
var
    Q : TQuery;
    Select, st1, st2 : string;
    i_ind1 : integer;
    TobEntete, TobLigne : TOB;
begin
//  la valeur a chercher est obligatoire
if CodeArticle = '' then Exit;
//  Initialisation de la Tob resultat
if TobResul = nil then
    TobResul := TOB.Create('', nil, -1)
else
    TobResul.ClearDetail;
//  Select en fonction de la colonne passée en parametre
Select := 'Select * from NOMENLIG where ';
if Colonne <> '' then
    Select := Select + Colonne + '="'
else
    Select := Select + 'GNL_CODEARTICLE="' + CodeArticle +
                       '" or GNL_SOUSNOMEN="';
Select := Select + CodeArticle + '"';
Q := OpenSQL(Select, True,-1, '', True);
if not Q.EOF then
    begin
    TobResul.LoadDetailDB('NOMENLIG', '', '', Q, False);
//  parcours des lignes pour renseigner le libellé des nomenclatures eventuelles
    TobEntete := TOB.Create('NOMENENT', nil, -1);
    for i_ind1 := 0 to TobResul.Detail.Count - 1 do
        begin
        TobEntete.PutValue('GNE_NOMENCLATURE', TobResul.Detail[i_ind1].GetValue('GNL_NOMENCLATURE'));
        if TobEntete.LoadDB then
            TobResul.Detail[i_ind1].PutValue('GNL_LIBELLE', TobEntete.GetValue('GNE_LIBELLE'));
//  si Complet, appel recursif sur la nomenclature courante pour verifier ses cas d'emploi
        if Complet then
            begin
            TobLigne := TobResul.Detail[i_ind1];
            CasEmploi(TobEntete.GetValue('GNE_NOMENCLATURE'), Colonne, Complet, TobLigne);
            end;
        st1 := TobResul.Detail[i_ind1].GetValue('GNL_NOMENCLATURE');
        st2 := TobResul.Detail[i_ind1].GetValue('GNL_SOUSNOMEN');
        TobResul.Detail[i_ind1].PutValue('GNL_SOUSNOMEN', st1);
        TobResul.Detail[i_ind1].PutValue('GNL_NOMENCLATURE', st2);
        TobResul.Detail[i_ind1].AddChampSupValeur ('LIBELLE', TobResul.Detail[i_ind1].GetValue('GNL_LIBELLE'));
{$IFDEF BTP}
        Tobresul.Detail[i_ind1].AddChampSup('QTEDUDETAIL', true);
        Tobresul.Detail[i_ind1].AddChampSup('UNITENOMENC', true);
        Tobresul.Detail[i_ind1].Putvalue('QTEDUDETAIL', 1);
        Tobresul.Detail[i_ind1].Putvalue('UNITENOMENC', '');
{$ENDIF}
        end;
    TobEntete.Free;
    end;
Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : B. Bicheray

Créé le ...... : 28/01/2000

Modifié le ... :   /  /

Description .. : Renseigne le champ supplémentaire LIBELLE

Mots clefs ... : NOMENCLATURE

*****************************************************************}

//=============================================================================
//  Renseigne le champ sup. LIBELLE avec 'CODEARTICLE (QUANTITE)'
//=============================================================================
procedure AffecteLibelle (var TobNomen : TOB; Colonne : string);
var TOBL : TOB ;
    i_ind1 : integer ;
    st1, st2, st3, st4 : string;
{$IFDEF BTP}
    st5,st6 : String;
{$ENDIF}
begin
for i_ind1 := 0 to TobNomen.Detail.Count - 1 do
    begin
    // MODIF BTP 19/06/03 if i_ind1 = 0 then TobNomen.Detail[i_ind1].AddChampSup('LIBELLE', True);
    st2 := VarAsType(TobNomen.Detail[i_ind1].GetValue(Colonne), varString);
    st4 := VarAsType(TobNomen.Detail[i_ind1].GetValue('GNL_SOUSNOMEN'), varString);
    st3 := VarAsType(TobNomen.Detail[i_ind1].GetValue('GNL_QTE'), varString);
{$IFDEF BTP}
    st5 := VarAsType(TobNomen.Detail[i_ind1].GetValue('QTEDUDETAIL'), varString);
    st6 := VarAsType(TobNomen.Detail[i_ind1].GetValue('UNITENOMENC'), varString);
    if st4 <> '' then
        begin
        st1 := st4 + ' (' + StrS0(Valeur(st3));
        if valeur(st5) > 0 then
             st1 := St1 + ' Détail pour '+ Strs0(valeur(st5))+ ' ' + st6 + ')'
        else
             st1 := St1 + ')';
        end else st1 := st2 + ' (' + StrS0(Valeur(st3)) + ')';
{$ELSE}
    if st4 <> '' then
        st1 := st4 + ' (' + StrS0(Valeur(st3)) + ')'
    else
        st1 := st2 + ' (' + StrS0(Valeur(st3)) + ')';
{$ENDIF}
    TobNomen.Detail[i_ind1].PutValue('LIBELLE', st1);
    if TobNomen.Detail[i_ind1].Detail.Count > 0 then
        begin
        TOBL := TobNomen.Detail[i_ind1];
        AffecteLibelle(TOBL, Colonne);
        end;
    end;
end;

//=============================================================================
//  Affichage du treeview
//=============================================================================
procedure AfficheTreeView(TV : TTreeView; TOBTemp : TOB; Colonne : string);
var
    st1 : string;
    i_ind1 : integer;
begin
    st1 := '';
    for i_ind1 := 0 to TOBTemp.MaxNiveau do
        st1 := st1 + ';' + Colonne;
    TOBTemp.PutTreeView(TV, nil, st1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : B. Bicheray

Créé le ...... : 28/01/2000

Modifié le ... :   /  /

Description .. : Recherche d'un noeud de treeview d'aprés son libellé

Mots clefs ... : NOMENCLATURE, TREEVIEW

*****************************************************************}

//-----------------------------------------------------------------------------
//  Recherche d'un noeud par son nom. Le point de depart est soit la
//  racine du treeview (TV <> nil, TN = nil), soit un noeud (TV = nil, TN <> nil)
//  Cette recherche peut se faire sur un ou plusieurs niveaux d'arborescence,
//  la valeur par defaut étant 1.
//-----------------------------------------------------------------------------
(*function ChercheNode(TV : TTreeView; TN : TTreeNode; Valeur : string; // DBR Fiche 10995
                     Niveau : integer = 1) : TTreeNode;
var
    st1 : string;
    tn1 : TTreeNode;
begin
Result := nil;
if TV <> nil then tn1 := TV.Items.GetFirstNode else tn1 := TN.getFirstChild;
while tn1 <> nil do
    begin
    if tn1.Level <= Niveau then
        begin
        st1 := tn1.Text;
        if Pos('(', tn1.Text) <> 0 then st1 := Copy(tn1.Text, 0, Pos('(', tn1.Text) - 2);
        if Valeur = st1 then
            begin
            Result := tn1;
            Exit;
            end;
        end;
    tn1 := tn1.GetNext;
    end;
end; *)

function ChercheNode(TV : TTreeView; TN : TTreeNode; Valeur : integer;
                     Niveau : integer = 1) : TTreeNode;
var
    tn1 : TTreeNode;
    i : integer ;
begin
Result := nil;
i:=0;
if TV <> nil then tn1 := TV.Items.GetFirstNode else tn1 := TN.getFirstChild;
while tn1 <> nil do
    begin
    if tn1.Level <= Niveau then
        begin
        i:=i+1;
        if i = valeur then
            begin
            Result := tn1;
            Exit;
            end;
        end;
    tn1 := tn1.GetNext;
    end;
end;
 // DBR Fiche 10995

//-----------------------------------------------------------------------------
//  Recherche du parent de niveau 1 d'un element du treeview.
//-----------------------------------------------------------------------------
function FindParentNiv1(TN : TTreeNode) : TTreeNode;
begin
if TN.Level = 1 then
    Result := TN
else
    Result := FindParentNiv1(TN.Parent);
end;

{***********A.G.L.***********************************************
Auteur  ...... : B. Bicheray

Créé le ...... : 07/02/2000

Modifié le ... :   /  /

Description .. : Recupére les libellés des dimensions et des grilles dimensions

Mots clefs ... : DIMENSION; ARTICLE

*****************************************************************}

procedure LibelleDimensions(CodeArticle : string; TOBArt : TOB; var LibelleDim, GrilleDim : string; Abrege : Boolean = False);
var
    iDim : integer;
    st1, Dimensions, RangDim, NomChamp : string;
    TobCode, TobDim : TOB;

begin
if (CodeArticle = '') and (TOBArt = nil) then Exit;
if CodeArticle = '' then CodeArticle := TOBArt.GetValue('GA_ARTICLE');
if TOBArt = nil then
    begin
    TobArt := TOB.Create('ARTICLE', nil, -1);
    if not TobArt.SelectDB('"' + CodeArticle + '"', nil, False) then Exit;
    end;
st1 := CodeArticleGenerique2(CodeArticle, Dimensions);
TobCode := TOB.Create('CHOIXCOD', nil, -1);
TobDim := TOB.Create('DIMENSION', nil, -1);
for iDim := 1 to MaxDimension do
    begin
    RangDim := ReadTokenSt(Dimensions) ;
    if Trim(RangDim) <> '' then
       BEGIN
       TobCode.PutValue('CC_TYPE', 'DIM');
       TobCode.PutValue('CC_CODE', 'DI' + IntToStr(iDim));
       TobCode.LoadDB;
       if Abrege then NomChamp := 'CC_ABREGE' else NomChamp := 'CC_LIBELLE' ;
       if iDim = 1 then
           LibelleDim := TobCode.GetValue(NomChamp)
           else
           LibelleDim := LibelleDim + ';' + TobCode.GetValue(NomChamp);
       TobDim.PutValue('GDI_TYPEDIM', 'DI' + IntToStr(iDim));
       TobDim.PutValue('GDI_GRILLEDIM', TobArt.GetValue('GA_GRILLEDIM' + IntToStr(iDim)));
       TobDim.PutValue('GDI_CODEDIM', RangDim);
       TobDim.ChargeCle1;
       TobDim.LoadDB;
       if iDim = 1 then
           GrilleDim := TobDim.GetValue('GDI_LIBELLE')
           else
           GrilleDim := GrilleDim + ';' + TobDim.GetValue('GDI_LIBELLE');
        END ;
    end;
TobCode.Free;
TobDim.Free;
end;

//
// Modif BTP
//
{$IFDEF BTP}

procedure RecupValorisation (TOBREF,TOBLIG,TOBMereLig : TOB);
var TOBCatalog    : TOB;
    TOBTarif      : TOB;
    TOBTiers      : TOB;
    MTPAF         : Double;
    PrixPourQte   : double;
    UA            : String;
    IsUniteAchat  : boolean;
begin
  //
  IsUniteAchat := true;
  //
  TOBTarif := TOB.Create ('TARIF',nil,-1);
  TOBCatalog := TOB.Create ('CATALOGU',nil,-1);
  TOBTiers := TOB.Create ('TIERS',nil,-1);
  TobTiers.putValue('T_TIERS',TOBRef.GetValue('GA_FOURNPRINC'));

  TRY
    ChargeTOBs (TOBRef,TOBCatalog,TOBTiers,TOBTarif);

    if TOBTarif.GetValue('GF_PRIXUNITAIRE') <> 0 then
      MTPAF :=TOBTarif.GetValue('GF_PRIXUNITAIRE')
    else
    begin
      prixPourQte := TOBCatalog.GetValue('GCA_PRIXPOURQTEAC');
      if PrixPourQte = 0 then PrixPourQte := 1;
      if TobCatalog.GetValue('GCA_PRIXBASE') <> 0 then
        MTPAF :=TOBCatalog.GetValue('GCA_PRIXBASE')/PrixPourQte
      else
        MTPAF :=TOBCatalog.GetValue('GCA_PRIXVENTE')/PrixPourQTe;
    end;
    UA := TOBCatalog.GetValue('GCA_QUALIFUNITEACH');

    //FV1 : 18/02/2015 - Gestion du Prix d'achat dans la fiche article
    if MTPAF = 0 then
    Begin
      MTPAF  := TOBRef.GetVAlue('GA_PAUA');
      UA     := TOBRef.GetValue('GA_QUALIFUNITEACH');
      IsUniteAchat := True;
      if MTPAF = 0 Then
      begin
        MTPAF := TOBRef.GetVAlue('GA_PAHT');
        IsUniteAchat := false;
      end;
    End;

    // Modif LS le 25/02/03
    if (TOBRef.GetValue('GA_TYPEARTICLE') <> 'OUV') then
    begin
      TOBLIG.putValue ('GCA_PRIXBASE',MTPAF);
      TOBLIG.putValue ('GA_FOURNPRINC',TOBRef.GetValue('GA_FOURNPRINC'));
      TOBLIG.putValue ('TGA_FOURNPRINC',TOBTIERS.GetValue('T_LIBELLE'));
      if TOBMereLig <> nil then
      begin
        if (TOBMereLig.GetValue('GA_TYPEARTICLE') = 'MAR')  or
          ((TOBMereLig.GetValue('GA_TYPEARTICLE') = 'ARP')  and
          (TOBLIG.getValue('GA_TYPEARTICLE')    <> 'PRE')) then
        begin
          TOBMereLig.putValue ('GCA_PRIXBASE',MTPAF);
          TOBMereLig.putValue ('GA_FOURNPRINC',TOBRef.GetValue('GA_FOURNPRINC'));
          TOBMereLIG.putValue ('TGA_FOURNPRINC',TOBTIERS.GetValue('T_LIBELLE'));
        end;
      end;
    end;

    MTPAF := arrondi(MTPAF * (1-(TOBTarif.GetValue('GF_REMISE')/100)),V_PGI.OkDecP );

    /// voila voila voila ..le seul hic c'est que ce prix est en UA..donc passage de l'UA en UV
    if IsUniteAchat then MTPAF := PassageUAUV (TOBRef,TOBCatalog,MTPAF);
    //
    if (MtPaf <> 0) and (mtpaf <> TobRef.getValue('GA_PAHT')) then
    begin
      TOBRef.putValue('GA_PAHT',MTPAF);
      RecalculPrPV (TOBRef,TOBCatalog);
    end;
    //
    if (TOBRef.GetValue('GA_TYPEARTICLE') = 'MAR') or (TOBRef.GetValue('GA_TYPEARTICLE')='ARP') or (TOBREF.GetValue('GA_TYPEARTICLE') = 'PRE') then
    begin
      TOBLIG.putValue('GA_PAHT',MTPAF);
      TOBLIG.putValue('GA_PVHT',TOBRef.GetValue('GA_PVHT'));
      //FV1 : 18/02/2015 - Gestion du Prix d'achat dans la fiche article
      TOBLIG.PutValue('GCA_QUALIFUNITEACH', UA);
      TOBLIG.putValue('GF_CALCULREMISE',TOBTarif.GetValue('GF_CALCULREMISE'));
      if TOBMereLig <> nil then
      begin
        if (TOBMereLig.GetValue('GA_TYPEARTICLE') = 'MAR') or
          ((TOBMereLig.GetValue('GA_TYPEARTICLE')='ARP') and (TOBLIG.getValue('GA_TYPEARTICLE') <> 'PRE')) then
        begin
          TOBMereLIG.putValue('GA_PAHT',MTPAF);
          TOBMereLIG.putValue('GA_PVHT',TOBRef.GetValue('GA_PVHT'));
          //FV1 : 18/02/2015 - Gestion du Prix d'achat dans la fiche article
          TOBMereLIG.PutValue('GCA_QUALIFUNITEACH', UA);
          //TOBMereLIG.putValue('GCA_QUALIFUNITEACH',TOBCatalog.GetValue('GCA_QUALIFUNITEACH'));
          TOBMereLIG.putValue('GF_CALCULREMISE',TOBTarif.GetValue('GF_CALCULREMISE'));
        end;
      end;
    end;
  FINALLY
    TOBTarif.free;
    TOBCatalog.free;
    TOBTiers.free;
  end;
end;

procedure RecupValoHrs (TOBRef,TOBMereLig : TOB);
var TOBPrestation : TOB;
    PVHT : double;
    QQ : TQuery;
    Req : String;
begin
if TOBMereLig = nil then Exit;
if TOBMereLig.getValue('GA_TYPEARTICLE') <> 'ARP' then exit;
TOBPrestation := TOB.Create ('ARTICLE',nil,-1);
TRY
   Req := 'SELECT * FROM ARTICLE WHERE GA_ARTICLE="'+TOBRef.GetValue('GA_ARTICLE')+'"';
   QQ := OpenSql (req,true,-1, '', True);
   TOBPrestation.selectDb ('',QQ);
   if not QQ.eof then
      begin
      if TOBPrestation.GetValue ('GA_ARTICLE') <> '' then
         begin
         if TOBPrestation.getValue ('GA_PRIXPOURQTE')=0 then TOBPrestation.PutValue('GA_PRIXPOURQTE',1);
      //
         PVHT := TOBPrestation.GetValue('GA_PVHT')/ TOBPrestation.getValue('GA_PRIXPOURQTE');
      // Réajustement au prix pour qte de l'article de base
         PVHT := PVHT * TOBRef.GetValue('GA_PRIXPOURQTE');

      // Récupération de la prestation associé
         TOBRef.putValue('TGA_PRESTATION',TobPrestation.GetValue('GA_LIBELLE')) ;
         //
         TOBMereLig.putvalue('TGA_PRESTATION',TobPrestation.GetValue('GA_LIBELLE')) ;
         TOBMereLig.putvalue('GA_NATUREPRES',TobRef.GetValue('GA_CODEARTICLE')) ;
         TOBMereLig.putvalue('GA_HEURE',TobRef.GetValue('GNL_QTE')) ;
         TOBMEreLig.PutValue('GF_PRIXUNITAIRE',PVHT);
         end;
      end;
   ferme (QQ);
FINALLY
    TOBPrestation.free;
END;
end;

{$ENDIF}

// Dans TOBIndice on a une ligne de nomenclature
procedure AppliqueValoDomaineAct (TOBIndice,TOBART : TOB);
var QQ : TQuery;
    TOBDOMACT : TOB;
begin
  if TOBIndice.getvalue('GNE_DOMAINE')='' then exit;
  TOBDOMACT := TOB.create ('BTDOMAINEACT',nil,-1);
  QQ := OpenSql('SELECT * FROM BTDOMAINEACT WHERE BTD_CODE="'+TOBIndice.getvalue('GNE_DOMAINE')+'"',true,-1, '', True);
  if not QQ.eof then
  begin
    TOBDOMACT.SelectDB ('',QQ);
    if TOBART.GetValue('GA_DPRAUTO')='X' then
    begin
      if TOBDOMACT.GetValue('BTD_COEFFG') <> 0 then
      begin
        TOBIndice.putValue('DPR',Arrondi(TOBIndice.getValue('DPA')*TOBDOMACT.GetValue('BTD_COEFFG'),V_PGI.okdecP));
        TOBIndice.putValue('COEFFG',TOBDOMACT.GetValue('BTD_COEFFG'));
      end;
    end;
    if TOBART.GetValue('GA_CALCAUTOHT')='X' then
    begin
      if TOBDOMACT.GetValue('BTD_COEFMARG') <> 0 then
      begin
        TOBIndice.putValue('PVHT',Arrondi(TOBIndice.getValue('DPR')*TOBDOMACT.GetValue('BTD_COEFMARG'),V_PGI.okdecP));
        TOBIndice.putValue('COEFMARG',TOBDOMACT.GetValue('BTD_COEFMARG'));
      end;
    end;
  end;
  ferme(QQ);
end;

procedure renseigneValoArt (TOBIndice,TOBART,TOBMereLig : TOB; SaisieComposant : Boolean=true);

  function CalculeQteFact (TOBL,TOBART : TOB) : double;
  begin
    result := TOBL.GeTDouble('GNL_QTESAIS');
    if not GetParamSocSecur('SO_BTGESTQTEAVANC',false) then exit;
    //
    if POS(TOBL.GetString('BNP_TYPERESSOURCE'),'SAL;INT;ST')> 0 then
    begin
      if TOBL.GetDouble('GNL_RENDEMENT')<> 0 then
      begin
        Result := Arrondi(TOBL.GeTDouble('GNL_QTESAIS')/ TOBL.GetDouble('GNL_RENDEMENT'),V_PGI.okdecQ);
      end else
      begin
        Result := TOBL.GetDouble('GNL_QTESAIS');
      end;
    end else
    begin
      if TOBL.GetDouble('GNL_PERTE')<> 0 then
      begin
        result := Arrondi(TOBL.GeTDouble('GNL_QTESAIS')* TOBL.GetDouble('GNL_PERTE'),V_PGI.okdecQ);
      end else
      begin
        Result := TOBL.GetDouble('GNL_QTESAIS');
      end;
    end;
  end;

var
   TOBREF : TOB;
   tcreated : boolean;
   Q : Tquery;
   Req : String;
begin
tcreated := false;
if TOBINdice.FieldExists ('GA_CALCPRIXPR') then
   begin
   TOBRef := TOBINDICE;
   end else
   begin
   if TOBArt <> nil then
      begin
      TOBRef := TOBArt;
      end else
      begin
      TOBRef := TOB.create ('ARTICLE',nil,-1);

      Q  := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
                     'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                     'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+
                     TOBIndice.getValue('GNL_ARTICLE')+'"',true,-1, '', True);
(*
      Req := 'SELECT GA_TYPEARTICLE,GA_ARTICLE,GA_CODEARTICLE,GA_FOURNPRINC,GA_TARIFARTICLE,GA2_SOUSFAMTARART,'
             + 'GA_CALCPRIXPR,GA_COEFFG,GA_CALCPRIXHT,GA_COEFCALCHT,'
             + 'GA_CALCPRIXTTC,GA_COEFCALCTTC,GA_DPR,GA_DPA,GA_PMAP,GA_PMRP,GA_PVHT,'
             + 'GA_PAHT,GA_PVTTC,GA_QUALIFUNITEVTE,GA_PRIXPOURQTE,GA_LIBCOMPL FROM ARTICLE'
             + ' LEFT JOIN ARTICLECOMPL ON GA2_ARTICLE=GA_ARTICLE';
      Req := Req + ' WHERE GA_ARTICLE = "' + TOBIndice.getValue('GNL_ARTICLE') + '"';
*)
//      Q := Opensql (req,true);
      if not Q.eof then
      begin
      	TOBRef.SelectDB ('',Q,true);
        InitChampsSupArticle (TOBRef);
      end else
      begin
      	TOBRef.InitValeurs;
        InitChampsSupArticle (TOBRef);
      end;
      ferme (Q);
      tcreated := true;
      end;
   end;
{$IFDEF BTP}
RecupValorisation (TOBREF,TOBIndice,TOBMereLig);
{$ENDIF}
// Récupération du Prix d'achat
{$IFDEF BTP}
TOBIndice.putvalue('DPA',TobRef.GetValue('GA_PAHT')) ;
TOBIndice.putvalue('COEFFG',TobRef.GetValue('GA_COEFFG')) ;
{$ELSE}
TOBIndice.putvalue('DPA',TobRef.GetValue('GA_DPA')) ;
{$ENDIF}

// Récupération du Prix de revient
TOBIndice.putvalue('COEFMARG',TobRef.GetValue('GA_COEFCALCHT')) ;
TOBIndice.putvalue('DPR',TobRef.GetValue('GA_DPR')) ;

TOBIndice.putvalue ('PVHT',valeur(strs(TobRef.GetValue('GA_PVHT'),V_PGI.okdecP)));
TOBIndice.putvalue ('PVTTC',valeur(strs(TobRef.GetValue('GA_PVTTC'),V_PGI.okdecP)));
TOBIndice.putvalue ('PMAP',valeur(strs(TobRef.GetValue('GA_PMAP'),V_PGI.okdecP)));
TOBIndice.putvalue ('PMRP',valeur(strs(TobRef.GetValue('GA_PMRP'),V_PGI.okdecP)));
//
(*
if TOBIndice.FieldExists('GA_TYPEARTICLE') then
begin
  TOBIndice.PutValue('GA_TYPEARTICLE',TobRef.GetValue('GA_TYPEARTICLE'));
end;
*)
TOBIndice.putvalue ('GA_QUALIFUNITEVTE',TobRef.GetValue('GA_QUALIFUNITEVTE'));
TOBIndice.putvalue ('GNL_PRIXPOURQTE',TobRef.GetValue('GA_PRIXPOURQTE'));
if TOBIndice.Getvalue ('GNL_PRIXPOURQTE') = 0 then TOBIndice.putvalue ('GNL_PRIXPOURQTE',1);
//
//TOBIndice.SetDouble('GNL_QTESAIS',TOBIndice.GetDouble('GNL_QTE'));
//
if (SaisieComposant) then TOBIndice.SetDouble('GNL_QTESAIS',TOBIndice.GetDouble('GNL_QTE'));
if (GetParamSocSecur('SO_BTGESTQTEAVANC',false)) and (SaisieComposant) then
begin
  //
  if (POS(TOBIndice.GetString('BNP_TYPERESSOURCE'),'SAL;INT;ST') > 0) then
  begin
    TOBIndice.PutValue('GNL_QUALIFUNITEACH',     TOBRef.GetValue('GA_QUALIFHEURE'));
    TOBIndice.PutValue('GNL_RENDEMENT',           TOBRef.GetValue('GA_COEFPROD'));
    TOBIndice.PutValue('GNL_PERTE',0);
  end else
  begin
    TOBIndice.SetString('GNL_QUALIFUNITEACH','');
    TOBIndice.PutValue('GNL_PERTE',           TOBRef.GetValue('GA_PERTEPROP'));
    TOBIndice.PutValue('GNL_RENDEMENT',0);
  end;
  TOBIndice.SetDouble('GNL_QTE',CalculeQteFact (TOBIndice,TOBRef));
end;

//TOBIndice.putvalue ('BNP_TYPERESSOURCE',TobRef.GetValue('BNP_TYPERESSOURCE'));
if (Pos(TOBIndice.GetValue('BNP_TYPERESSOURCE'),VH_GC.BTTypeMoInterne)>0) then
begin
  TOBIndice.putvalue ('GA_HEURE',TOBIndice.GetValue('GNL_QTE'));
  TOBIndice.putvalue ('GF_PRIXUNITAIRE',TOBIndice.GetValue('PVHT'));
  TOBIndice.putvalue ('GA_PVHT',0);
  TOBIndice.putvalue ('TPSUNITAIRE',1);
end else TOBIndice.putvalue ('TPSUNITAIRE',0);
// Dans le cas ou un domaine d'activiet est renseigné on applique sur les articles de type marchandise les coef et
// on recalcule les DPR et PUHT
if (TOBIndice.GetValue('GNE_DOMAINE')<>'') and (TOBRef.getValue('GA_PRIXPASMODIF')<>'X') then
begin
  AppliqueValoDomaineAct (TOBIndice,TOBRef);
  TOBIndice.putValue('GA_PAHT',TOBIndice.getValue('DPA'));
  TOBIndice.putValue('GA_PVHT',TOBIndice.getValue('PVHT'));
  TOBIndice.putvalue ('GF_PRIXUNITAIRE',TOBIndice.GetValue('PVHT'));
end;

//
if tcreated then TOBREf.free;
end;

procedure ValoriseOuvrage (CodeNomen : string;TOBMereEnt,TOBMereLig : TOB; var Resul : T_Valeurs;AvecPourcent:boolean=true;RecupEntete:boolean=false);
var
   TOBEnt : TOB;
   tcreated : boolean;
begin
tcreated := false;
//  init du tableaux résultat
InitTableau (Resul);
if (CodeNomen = '') then Exit;
if (RecupEntete) then
   begin
   if TOBMereEnt <> nil then
      TOBEnt := TOBMereEnt
   else
       begin
       ChargeEntete (CodeNomen,TobEnt);
       tcreated := true;
       end;
   end else TOBEnt := TOBMereEnt;
CalculeOuvrage (CodeNomen,TOBEnt,TOBMereLig,ExAucun,Resul,AvecPourcent);
if tcreated then TOBEnt.free;
end;


procedure CalculeOuvrage (CodeNomen : String; TObMereEntete,TOBMereLigne : TOB; Exception: T_ExceptVal ;var Resul : T_Valeurs;AvecPourcent:boolean=true);
var
   Select: String;
   Q: Tquery;
   TobEnt,TOBdetailLig,TOBIndice : Tob;
   Indice : Integer;
   IPercent : Integer;
   Tlocal,TSumPourcent : T_Valeurs;
   ArticleOk : string;
begin
  initTableau (TSumPourcent);
  IPercent := -1;
  // chargement du détail ouvrage
  Select := 'Select NOMENLIG.*,GNE_DOMAINE,GA_PRIXPASMODIF,GA_TYPEARTICLE,GA_ARTICLE,GA_CODEARTICLE,GA_FOURNPRINC,GA_QUALIFUNITEVTE,'
            + 'GA_TARIFARTICLE,GA2_SOUSFAMTARART,GA_PRIXPOURQTE,GA_CALCPRIXPR,GA_DPRAUTO,GA_COEFFG,GA_LIBCOMPL,'
            + 'GA_CALCAUTOHT,GA_COEFCALCHT,GA_COEFCALCTTC,GA_CALCPRIXHT,GA_CALCPRIXTTC,GA_COEFCONVQTEVTE,'
            + 'GA_DPA,GA_DPR,GA_PMAP,GA_PMRP,GA_PAHT,GA_PAUA,GA_PVHT,GA_PVTTC FROM NOMENLIG '
            + 'LEFT JOIN ARTICLE ON GA_ARTICLE=GNL_ARTICLE LEFT JOIN ARTICLECOMPL ON GA2_ARTICLE=GNL_ARTICLE '
            + 'LEFT JOIN NOMENENT ON GNE_NOMENCLATURE=GNL_NOMENCLATURE '
            + 'where GNL_NOMENCLATURE = "' + CodeNomen + '"';
  Q := OpenSQL(Select, true,-1, '', True);
  if not Q.EOF then
  begin
    // C'est bien une nomenclature
    TOBdetailLig := TOB.create ('LNOMENLIG',nil,-1);
    TOBdetailLig.LoadDetailDB('DNOMENLIG', '', '', Q, False);
    ferme(Q);
    TOBIndice := TOBDetailLig.FindFirst (['GA_TYPEARTICLE'],['POU'],false);
    if TOBIndice <> nil then ArticleOk := TOBIndice.getVAlue('GA_LIBCOMPL');
    for indice:= 0 to TOBdetailLig.Detail.Count -1 do
    begin
      TOBIndice := TOBdetailLig.detail[indice];
      ChampSupNomen (TOBIndice);
      If TOBIndice.GetValue ('GA_PRIXPOURQTE') = 0 then TOBIndice.PutValue ('GA_PRIXPOURQTE',1);
      if TobIndice.GetValue('GNL_SOUSNOMEN') <> '' then
      begin
        ChargeEntete (CodeNomen,TobEnt);
        initTableau (Tlocal);
        CalculeOuvrage (TOBIndice.getvalue ('GNL_SOUSNOMEN'),TobEnt,nil, Exception,TLocal);
        TOBIndice.putvalue ('DPA',Tlocal[0]);
        TOBIndice.putvalue ('DPR',Tlocal[1]);
        TOBIndice.putvalue ('PVHT',Tlocal[2]);
        TOBIndice.putvalue ('PVTTC',Tlocal[3]);
        TOBIndice.putvalue ('PMAP',Tlocal[6]);
        TOBIndice.putvalue ('PMRP',Tlocal[7]);
        CumuleElementTable (TOBIndice,Exception,resul);
      end else
      begin
        renseigneValoArt (TOBIndice,nil,TOBMereLigne);
        {$IFDEF BTP}
        if (TOBIndice.getValue('GA_TYPEARTICLE') = 'PRE') then RecupValoHrs (TOBIndice,TOBMereLigne);
        {$ENDIF}
        if TobIndice.GetValue('GNL_PRIXFIXE') <> 0 then
        TOBIndice.putvalue ('PVHT',TOBIndice.getValue('GNL_PRIXFIXE'));
        // On est sur une ligne
        if TOBIndice.GetValue ('GA_TYPEARTICLE') = 'POU' then
        begin
          Ipercent := Indice;
          continue;
        end;
        CumuleElementTable (TOBIndice,Exception,resul);
        if ArticleOKInPOUR (TOBIndice.Getvalue('GA_TYPEARTICLE'),ArticleOk) then CumuleElementTable (TOBIndice,Exception,TSumPourcent);
      end;
      FormatageTableau (resul,V_PGI.okdecP);
      if (IPercent >= 0) and (AvecPourcent=true) then
      begin
        Tlocal := CalculSurTableau ('*',TSumPourcent,TOBdetailLig.detail[IPercent].getvalue('GNL_QTE')/valeur('100'));
        Tlocal[2] := arrondi(TLocal[2],V_PGI.OkDecP );
        FormatageTableau (Tlocal,V_PGI.okdecP);
        resul := CalculSurTableau ('+',resul,Tlocal);
        FormatageTableau (resul,V_PGI.okdecP);
      end;
    end;
    if TOBMereEntete <> nil then
    begin
      if TOBMereEntete.getValue ('GNE_QTEDUDETAIL') = 0 then TOBMereEntete.putValue('GNE_QTEDUDETAIL',1);
      resul := CalculSurTableau ('/',resul,TobMereEntete.getValue('GNE_QTEDUDETAIL'));
      FormatageTableau (resul,V_PGI.okdecP);
    end;
    TOBdetailLig.free;
  end else
  begin
    Ferme(Q);
  end;
end;

procedure ChargeEntete (CodeNomen : String; var TOBEnt : Tob);
var
   Q,Qart : Tquery;
   Select : String;
   TobArt : TOB;
begin
     TobEnt := TOB.create ('NOMENENT',nil,-1);
     TobEnt.AddChampSup ('GNE_QTEDUPV',false);
     TobEnt.PutValue ('GNE_QTEDUPV',valeur('1'));
     TobEnt.PutValue ('GNE_QTEDUDETAIL',valeur('1'));

     Select := 'Select * from NOMENENT where GNE_NOMENCLATURE = "' + CodeNomen + '"';
     Q := OpenSQL(Select, False,-1, '', True);
     if not Q.EOF then
     begin
        TobEnt.SelectDb('',Q,false);
        TobArt := TOB.create ('ARTICLE',nil,-1);
        Select := 'Select GA_PRIXPOURQTE from ARTICLE where GA_ARTICLE = "' + TobEnt.getvalue('GNE_ARTICLE') + '"';
        QArt := OpenSQL(Select, true,-1, '', True);
        if not QArt.eof then
        begin
             if TobArt.getValue('GA_PRIXPOURQTE') <> 0 then
                TobEnt.putValue ('GNE_QTEDUPV',TobArt.getvalue('GA_PRIXPOURQTE'));
        end;
        ferme (Qart);
        TOBArt.free;
     end;
     ferme (Q);
end;

procedure CumuleElementTable (TOBDetail : TOB ; Exception : T_ExceptVal; var resul:T_Valeurs);
Var Qtedudetail : Integer;
begin
  QteDuDetail := TOBDetail.getValue('GNL_PRIXPOURQTE') * TOBDetail.getValue('QTEDUDETAIL');
  if QteDuDetail = 0 then QteDuDetail := 1;

  if (ExAucun = Exception) or (ExPVHt = Exception) then
  begin
    // Récupération du Prix d'achat
    Resul[0] := Resul[0] + Arrondi((TobDetail.GetValue('DPA')* (TObDetail.Getvalue('GNL_QTE')/ QteDuDetail)),V_PGI.OkdecV);
    Resul[1] := Resul[1] + Arrondi((TobDetail.GetValue('DPR')* (TObDetail.Getvalue('GNL_QTE')/ QteDuDetail)),V_PGI.OKdecV);
    Resul[3] := Resul[3] + Arrondi((TobDetail.GetValue('PVTTC')* (TObDetail.Getvalue('GNL_QTE') / QteDuDetail)),V_PGI.OkdecV);
    Resul[6] := Resul[6] + Arrondi((TobDetail.GetValue('PMAP')* (TObDetail.Getvalue('GNL_QTE')/ QteDuDetail)),V_PGI.OkdecV);
    Resul[7] := Resul[7] + Arrondi((TobDetail.GetValue('PMRP')* (TObDetail.Getvalue('GNL_QTE')/ QteDuDetail)),V_PGI.OkdecV);
    if (Pos(TobDetail.GetValue('BNP_TYPERESSOURCE'),VH_GC.BTTypeMoInterne)>0) or
       (TobDetail.Detail.count > 0) then
      Resul[9] := Resul[9] +  (TobDetail.GetValue('TPSUNITAIRE')* TObDetail.Getvalue('GNL_QTE'));
  end;
  if (ExPVUniq = Exception) or (ExAucun = Exception) then
  begin
  if TOBDetail.getValue('GNL_PRIXFIXE') <> 0 then
  begin
    Resul[2] := Resul[2] + Arrondi((TobDetail.GetValue('GNL_PRIXFIXE')* (TObDetail.Getvalue('GNL_QTE') / QteDuDetail)),V_PGI.OkdecV);
    if TobDetail.FieldExists('GCA_PRIXVENTE') then TobDetail.PutValue('GCA_PRIXVENTE',Arrondi(TobDetail.GetValue('GNL_PRIXFIXE')* (TObDetail.Getvalue('GNL_QTE')),V_PGI.okdecV));
  end else
  begin
    Resul[2] := Resul[2] + Arrondi((TobDetail.GetValue('PVHT')* (TObDetail.Getvalue('GNL_QTE') / QteDuDetail)),V_PGI.OkdecV);
    if TobDetail.FieldExists('GCA_PRIXVENTE') then TobDetail.PutValue('GCA_PRIXVENTE',Arrondi(TobDetail.GetValue('PVHT')* (TObDetail.Getvalue('GNL_QTE')),V_PGI.okdecV));
  end;
  end;
end;

Function UniteOuvrage (article : string): String;
var
   Q:Tquery;
   Req : String;
begin
   Req := 'SELECT GA_QUALIFUNITEVTE FROM ARTICLE WHERE GA_ARTICLE="' + article + '"';
   Q := opensql (Req,true,-1, '', True);
   if not Q.eof then
      Result := Q.FindField('GA_QUALIFUNITEVTE').AsString
   else
      Result := '';
   Ferme (Q);
end;

procedure FormatageTableau (var resul : T_valeurs;Decimale:integer=-1);
var indice : Integer;
begin
   for Indice := 0 to length(resul)-1 do
   begin
   	 if Indice=9 then continue;	
     if decimale = -1 then Resul[Indice] := arrondi (Resul[Indice], V_PGI.OkDecP)
                      else Resul[Indice] := arrondi (Resul[Indice], Decimale);
   end;
end;

procedure InitTableau (var resul : T_valeurs);
var indice : Integer;
begin
   for Indice := 0 to length(resul)-1 do
   begin
     Resul[Indice] := valeur('0');
   end;
end;

function CalculSurTableau ( operation : String; entree : T_valeurs; operande : T_valeurs): T_valeurs;
var
   TLocal : T_Valeurs;
   Indice : Integer;
begin
for Indice := 0 to length(entree)-1 do
    begin
    TLocal [Indice ] := valeur ('0');
    if operation = '+' then
       begin
       TLocal [Indice]  := entree [Indice] + operande [Indice];
       end
    else if operation = '-' then
       begin
       TLocal [Indice]  := entree [Indice] - operande [Indice];
       end
    else if operation = '*' then
       begin
       TLocal [Indice]  := entree [Indice] * operande [Indice];
       end
    else if operation = '/' then
       begin
       if operande[indice] <> 0 then TLocal [Indice]  := entree [Indice] / operande [Indice];
       end;
end;
result := TLocal;
end;

function CalculSurTableau ( operation : String; entree : T_valeurs; operande : double): T_valeurs;
var
   TLocal : T_Valeurs;
   Indice : Integer;
begin
for Indice := 0 to length(entree)-1 do
    begin
    TLocal [Indice ] := valeur ('0');
    if operation = '+' then
       begin
       TLocal [Indice]  := entree [Indice] + operande;
       end
       else if operation = '-' then
       begin
       TLocal [Indice]  := entree [Indice] - operande;
       end
       else if operation = '*' then
       begin
       TLocal [Indice]  := entree [Indice] * operande;
       end
       else if operation = '/' then
       begin
       if operande <> 0 then TLocal [Indice]  := entree [Indice] / operande;
       end;
end;
result := TLocal;
end;
// --------------------------------

procedure RepriseDonneeArticle (TOBLN,TOBArticle : TOB; FromCopierColler : boolean = false; WithValorisation : boolean= true);
{$IFDEF BTP}
var TOBL : TOB;
    Indice : Integer;
    Q : Tquery;
    TOBA : TOB;
{$ENDIF}
begin
{$IFDEF BTP}
// Mise en place des valeurs achats
if TOBLN.detail.count > 0 then
   begin
   for Indice := 0 to TOBLN.detail.count -1 do
       begin
       TOBL := TOBLN.detail[Indice];
       Q := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
                     'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                     'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_CODEARTICLE="'+
                      TOBL.GetValue('BLO_CODEARTICLE') + '" AND GA_STATUTART <> "DIM" ',true,-1, '', True);
       TOBA := TOB.create ('ARTICLE',nil,-1);
       TOBA.selectdb ('',q);
       InitChampsSupArticle (TOBA);
       ferme (Q);
       RepriseDonneeArticle (TOBL,TOBA,FromCopierColler,WithValorisation);
       TOBA.free;
       end;
// Modif LS le 16/01/03
   if TOBLN.fieldExists ('ANCPV') then TOBLN.putvalue('ANCPV',valeur(TOBLN.GetValue('ANCPV')));
   if TOBLN.fieldExists ('ANCPA') then TOBLN.putvalue('ANCPA',valeur(TOBLN.GetValue('ANCPA')));
   if TOBLN.fieldExists ('ANCPR') then TOBLN.putvalue('ANCPR',valeur(TOBLN.GetValue('ANCPR')));
// --
   end else
   begin
   if not TOBLN.FieldExists ('BLO_CODEARTICLE') then exit;
   if TobArticle = nil then
      begin
       Q := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
                     'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                     'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_CODEARTICLE="'+
                     TOBLN.GetValue('BLO_CODEARTICLE') + '" AND GA_STATUTART <> "DIM" ',true,-1, '', True);
       TOBA := TOB.create ('ARTICLE',nil,-1);
       TOBA.selectdb ('',q);
       InitChampsSupArticle (TOBA);
       ferme (Q);
      end else TOBA := TOBARTICLE;
   if not FromCopierColler then
   begin
     if TOBLN.GetValue('BLO_FOURNISSEUR')=''  then
     begin
       if WithValorisation then
       begin
         if (VH_GC.ModeValoPa = 'PMA') and (pos(TOBA.getString('GA_TYPEARTICLE'),'MAR;ARP')>0) then
         begin
           if TOBA.GetDouble('GA_PMAP') <> 0 then
           begin
         		TobLN.PutValue('BLO_DPA',TOBA.GetValue('GA_PMAP'));
           end else
           begin
           	TobLN.PutValue('BLO_DPA',TOBA.GetValue('GA_PAHT'));
           end;
         end else
         begin
         	TobLN.PutValue('BLO_DPA',TOBA.GetValue('GA_PAHT'));
         end;
         TobLN.PutValue('BLO_DPR',TOBA.GetValue('GA_DPR'));
         TobLN.PutValue('BLO_PMAP',TOBA.GetValue('GA_PMAP'));
         TobLN.PutValue('BLO_PMRP',TOBA.GetValue('GA_PMRP'));
       end;
     end;
     TobLN.PutValue('BLO_FAMILLETAXE1',TOBA.GetValue('GA_FAMILLETAXE1'));
     TobLN.PutValue('BLO_FAMILLETAXE2',TOBA.GetValue('GA_FAMILLETAXE2'));
     TobLN.PutValue('BLO_FAMILLETAXE3',TOBA.GetValue('GA_FAMILLETAXE3'));
     TobLN.PutValue('BLO_FAMILLETAXE4',TOBA.GetValue('GA_FAMILLETAXE4'));
     TobLN.PutValue('BLO_FAMILLETAXE5',TOBA.GetValue('GA_FAMILLETAXE5'));
   end;
   TOBLN.putvalue('ANCPV',valeur(TOBLN.GetValue('ANCPV')));
   TOBLN.putvalue('ANCPA',valeur(TOBLN.GetValue('ANCPA')));
   TOBLN.putvalue('ANCPR',valeur(TOBLN.GetValue('ANCPR')));
   GetValoDetail (TOBLN);
   if TobArticle = nil then TOBA.free;
   end;
{$ENDIF}
end;


procedure ReajusteLigneParDoc (TOBLD,TOBLN: TOB;DEV:RDevise;Niveau : integer);
{$IFDEF BTP}
var Indice : Integer;
    TOBl,TOBPiece : TOB;
    X,Taux,NewTaux,Quotite,NewQuotite,XP,XC : double;
    Valeurs:T_valeurs;
    GestionHt : boolean;
{$ENDIF}
begin
  {$IFDEF BTP}
  GestionHt := (TOBLD.GetValue('GL_FACTUREHT')='X');
  TOBPiece := TOBLD.Parent ;
  if TOBLN.detail.count > 0 then
  begin
    for Indice := 0 to TOBLN.detail.count -1 do
    begin
      TOBL := TOBLN.detail[Indice];
      ReajusteLigneParDoc (TOBLD,TOBL,DEV,Niveau+1);
    end;
    CalculeOuvrageDoc (TOBLN,1,1,true,DEV,Valeurs,GestionHt,(TOBLD.GetValue('GL_BLOQUETARIF')='-'));
    TOBLN.Putvalue ('BLO_DPA',Valeurs[0]);
    TOBLN.Putvalue ('BLO_DPR',Valeurs[1]);
    if TOBLD.GetValue('GL_BLOQUETARIF') = '-' then
    begin
      TOBLN.Putvalue ('BLO_PUHTDEV',Valeurs[2]);
      TOBLN.Putvalue ('BLO_PUTTCDEV',Valeurs[3]);
    end;
    TOBLN.Putvalue ('BLO_PMAP',Valeurs[6]);
    TOBLN.Putvalue ('BLO_PMRP',Valeurs[7]);
    TOBLN.Putvalue ('BLO_TPSUNITAIRE',Valeurs[9]);
    StockeLesTypes (TOBLN,Valeurs);
  end else
  begin
     Taux := TOBLN.GetValue('BLO_TAUXDEV');
     Quotite := TOBLN.GetValue('BLO_COTATION');
     NewTaux := TOBLD.GetValue('GL_TAUXDEV');
     NewQuotite := TOBLD.GetValue('GL_COTATION');
     X := TOBLN.GetValue ('BLO_PUHTDEV');
     if TOBLN.GetValue('BLO_DEVISE') <> TOBLD.GetValue('GL_DEVISE') then
        X := DEVISETODEVISE (X,Taux,Quotite,NewTaux,NewQuotite,V_PGI.OkdecP );
     TOBLN.putvalue ('BLO_PUHTDEV',X);
     X := TOBLN.GetValue ('BLO_PUTTCDEV');
     if TOBLN.GetValue('BLO_DEVISE') <> TOBLD.GetValue('GL_DEVISE') then
        X := DEVISETODEVISE (X,Taux,Quotite,NewTaux,NewQuotite,V_PGI.OkdecP);
     TOBLN.putvalue ('BLO_PUTTCDEV',X);
  end;
  if GestionHt then TOBLN.PutValue ('ANCPV', TOBLN.getvalue('BLO_PUHTDEV'))
              else TOBLN.PutValue ('ANCPV', TOBLN.getvalue('BLO_PUTTCDEV'));
  TOBLN.PutValue ('ANCPA', TOBLN.getvalue('BLO_DPA'));
  TOBLN.PutValue ('ANCPR', TOBLN.getvalue('BLO_DPR'));
  TOBLN.PutValue ('BLO_DEVISE', TOBLD.getvalue('GL_DEVISE'));
  TOBLN.PutValue ('BLO_TAUXDEV', TOBLD.getvalue('GL_TAUXDEV'));
  TOBLN.PutValue ('BLO_COTATION', TOBLD.getvalue('GL_COTATION'));
  TOBLN.PutValue ('BLO_REGIMETAXE', TOBLD.getvalue('GL_REGIMETAXE'));
  TOBLN.PutValue ('BLO_FACTUREHT', TOBLD.getvalue('GL_FACTUREHT'));
  TOBLN.PutValue ('BLO_TVAENCAISSEMENT', TOBLD.getvalue('GL_TVAENCAISSEMENT'));
  TOBLN.PutValue ('BLO_AFFAIRE', TOBLD.getvalue('GL_AFFAIRE'));
  TOBLN.PutValue ('BLO_AFFAIRE1', TOBLD.getvalue('GL_AFFAIRE1'));
  TOBLN.PutValue ('BLO_AFFAIRE2', TOBLD.getvalue('GL_AFFAIRE2'));
  TOBLN.PutValue ('BLO_AFFAIRE3', TOBLD.getvalue('GL_AFFAIRE3'));
  TOBLN.PutValue ('BLO_AVENANT', TOBLD.getvalue('GL_AVENANT'));
  TOBLN.PutValue ('BLO_REPRESENTANT', TOBLD.getvalue('GL_REPRESENTANT'));
  TOBLN.PutValue ('BLO_APPORTEUR', TOBLD.getvalue('GL_APPORTEUR'));
  TOBLN.PutValue ('BLO_COMMISSIONR', TOBLD.getvalue('GL_COMMISSIONR'));
  TOBLN.PutValue ('BLO_COMMISSIONA', TOBLD.getvalue('GL_COMMISSIONA'));
  TOBLN.PutValue ('BLO_TYPECOM', TOBLD.getvalue('GL_TYPECOM'));
  TOBLN.PutValue ('BLO_VALIDECOM', TOBLD.getvalue('GL_VALIDECOM'));
  TOBLN.PutValue ('BLO_DEPOT', TOBLD.getvalue('GL_DEPOT'));
  TOBLN.PutValue ('BLO_FACTUREHT', TOBLD.getvalue('GL_FACTUREHT'));
  TOBLN.PutValue('BLO_SOUCHE',TOBLD.GetValue('GL_SOUCHE')) ;
  TOBLN.PutValue('BLO_NUMERO',TOBLD.GetValue('GL_NUMERO')) ;
  TOBLN.PutValue('BLO_INDICEG',TOBLD.GetValue('GL_INDICEG')) ;
  TOBLN.PutValue('BLO_NUMLIGNE',TOBLD.GetValue('GL_NUMLIGNE')) ;
  TOBLN.PutValue('BLO_NIVEAU',Niveau) ;
  if GestionHt then
  BEGIN
    CalculeLigneHTOuv  (TOBLN,TOBPiece,DEV);
  END ELSE
  BEGIN
    CalculeLigneTTCOuv  (TOBLN,TOBPiece,DEV);
  END;
  CalculeMontantsAssocie (TOBLN.GetValue('BLO_PUHTDEV'),XP,DEV);
  TOBLN.PutValue('BLO_PUHT',XP);
  CalculeMontantsAssocie (TOBLN.GetValue('BLO_PUTTCDEV'),XP,DEV);
  TOBLN.PutValue('BLO_PUTTC',XP);
  if not VariantePieceAutorisee (TOBPiece) then TOBLN.PutValue('BLO_TYPELIGNE','ART');
  {$ENDIF}
end;

{$IFDEF BTP}
function ExistNomen (Article,Nomen : string; TOBREF : TOB = nil) : boolean;
var ThisTOb : TOB;
begin
  result := false;
  if TOBREF <> nil then
  begin
    ThisTOB := TOBREF.FindFirst (['GNE_NOMENCLATURE'],[Nomen],true);
    if ThisTOB <> nil then result := true;
  end;
  if not result then result := ExisteSQL('Select GNE_NOMENCLATURE from NOMENENT where GNE_NOMENCLATURE="'+Nomen+'"');
end;

function AffecteCodeNomen (CodeArticle,CodeNomenRef : string; TOBREF : TOB = nil) : string;
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
      ok := ExistNomen (CodeArticle,NewCode,TOBREF);
      if ok then inc(Indice) else result := NewCode;
    until not ok;
  end else
  begin
    repeat
      if length(inttostr(Indice))> 10 then break;
      NewCode := copy (CodeNomenRef,1,35-length(inttostr(Indice)))+inttostr(Indice);
      ok := ExistNomen (CodeArticle,NewCode,TOBREF);
      if ok then inc(Indice) else result := NewCode;
    until not ok;
  end;
end;

// controle si l'article est decliné
function CompteDeclinaison(Article: string): integer;
var requete : string;
    QQ: TQuery;
begin
      Requete := 'SELECT COUNT(*) FROM NOMENENT WHERE GNE_ARTICLE="' + Article + '"';
      QQ := Opensql(Requete, True,-1, '', True);
      Result := QQ.Fields[0].AsInteger;
      Ferme(QQ);
end;

{$ENDIF}
end.
