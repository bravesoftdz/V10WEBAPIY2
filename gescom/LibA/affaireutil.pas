unit AffaireUtil;
                                                        
// mcd 28/03/03 ce source à été revu pour sortir des fct de celui-ci
//CodeAffaireAffiche, CodeaffaireDecoupe, OleFormatexercice,
//AfficheCOdeAff,Libexercice,TeststatuReduit,TestCodeAvenant
// afin de les mettre dans la source CalcOleGenericAff
// pour utilisation dans les @ des états
interface               
Uses
    Classes,HEnt1,Paramsoc,ComCtrls,sysutils,HCtrls,
  EntGC,hmsgbox,UTob,FactComm,FactUtil,uEntCommun,
  {$IFNDEF ERADIO}
    LookUp,
  {$ENDIF !ERADIO}
  {$IFNDEF EAGLSERVER}
    AGLInit,
  {$ELSE  !EAGLSERVER}
    uUtilProcessServer,
  {$ENDIF !EAGLSERVER}
    Controls,Formule,UtilArticle,tiersutil,
{$IFDEF EAGLCLIENT}
     MaineAGL,emul,
{$ELSE}
    {$IFNDEF ERADIO}
      Fe_Main,
    {$ENDIF !ERADIO}
    {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB, Mul,UtilGc,
{$ENDIF}
     DicoBTP,UtilPGI,HTB97,Forms,SaisUtil,Graphics,
     AfUtilArticle,Ent1,UtilRessource ,
{$IFDEF BTP}
     CalcOleGenericBTP, BTPUtil, AppelsUtil,
{$ENDIF}
     FactTOB,
     FactAdresse,UtilTOBPiece ;
     
Type
  T_OP = Class (TObject)
    DateCourante:TDateTime;
    TypeBlocage:string;
    tobAffaire:TOB;
    function DonneesOperation(AcsChampFormule:Hstring):variant;
    end ;


Const

  NbMaxPartieAffaire =3;   // mcd 28/03/03 si modif voir valeur en dur dans source CalcOleGenericAff
  cInClient       = 0; // standard pas de numero de client
  cInClientAlgoe  = 1; // algoe
  cInClientOnyx   = 2; // onyx
  cInClientAmyot  = 3; // Amyot

Var DicoAffaire : array[1..2,1..20] of string ;

      TexteMsgPieceAffaire: array[1..4] of string 	= (
      {1}        'Code affaire invalide',
      {2}        'Plus d''une pièce associée à l''affaire',
      {3}        'Aucune pièce associée à l''affaire',
      {4}        'Code Tiers non valide'
                 );
// OPTIMISATIONS LS
    TOBlesAffaires : TOB;
// ------------
Type T_TypeModifAff = Set of (tmaMntEch,tmaPourcEch,tmaDate,tmaDFacLiq,tmaRepAct) ;
Type T_TypeBlocAff = (tbaAucun, tbaDates, tbaEtat, tbaConfident) ;


// Fonction de chargement et de déchargement du code affaire à partir des composants
function ChargeCleAffaire(Part0, Part1,Part2,Part3,Avenant: THCritMaskEDIT; BRechAffaire:TToolBarButton97; FTypeAction:TActionFiche; CodeAff:string;SaisieAffaire:Boolean) : integer;
Procedure CalculeCodeAffaire(Part0,Part1,Part2,Part3,Avenant: THCritMaskEDIT;BRechAffaire:TToolBarButton97; FTypeAction:TActionFiche; CodeAff:string;SaisieAffaire:Boolean;FromAppel:boolean=false);
function DechargeCleAffaire(Part0, Part1,Part2,Part3,Avenant: THCritMaskEDIT; CodeTiers : string;
                              FTypeAction:TActionFiche; SaisieAffaire,AfficheMsg,bProposition:Boolean;
                              var AviPartErreur:integer):String; // Modif PL le 15/06/2000 : ajout de AviPartErreur pour savoir quelle zone sort en erreur

// Fonction de chargement et de déchargement du code affaire à partir de string
function  CodeAffaireRegroupe(Var Part0, Part1 ,Part2 ,Part3, Avenant : string ; FTypeAction:TActionFiche ; SaisieAffaire,bGarderCpt,bUpdateCpt : Boolean): String;
// gestion passage code affaire complet à celui affiché
Function MaxLngAffaireAffiche : integer;
Function CodeAffAfficheToComplet (CodeAffReduit,Tiers : string) : string;
Function CodeAffaireUser (CodeAffaire : String) : String;

 // fct incr/decr zone exercice (mcd 02/10/2002 pris de affaireduplic)
function  CodificationNewExercice (Exercice : string;Plus : Boolean): string ;

function  CodeAffaireSuivant ( Code : String ; Lng : Integer) : String ;
function  ChercheMissionExercice (StatutReduit, Partie1 , Partie2, Avenant, Codetiers : string) : Boolean ;
{$IFDEF EAGLCLIENT}
Procedure ModifColAffaireGrid (Gr : THGrid; TobQ : TOB );
{$ELSE}
  {$IFNDEF ERADIO}
    Procedure ModifColAffaireGrid (Gr : TfMul );
    Procedure FormatExerciceGrid  (Gr : THDBGrid );
  {$ENDIF !ERADIO}
{$ENDIF}


Function  RegroupePartiesAffaire (Var Part0,Part1, Part2, Part3,Avenant : String) : string;
Function  CompteurPartieAffaire  ( NumPartie : Integer; bProposition,bUpdateCpt : boolean; FromAppel:boolean=false ) : string;
function  LookUpPartiesAffaire   ( NumPartie : Integer ; GS : THGrid ) : boolean;

Function  TestCleAffaire(AFFAIRE,AFFAIRE1,AFFAIRE2,AFFAIRE3,AVENANT,TIERS:THCritMaskEdit;Part0 : string; bAvecMessage, bAvecAffModele, bProposition:boolean):integer;
Function TeststCleAffaire(var Affaire,Affaire0,Affaire1,Affaire2,Affaire3,Avenant,Tiers : string; bAvecMessage, bAvecAffModele,bProposition,WhereParties:boolean):integer;
Procedure GoSuiteCodeAffaire(PartEnc,PartSuite : THCritMaskEDIT; NumPartieEnc: Integer);
Function  ChargeAffaire(AFFAIRE,AFFAIRE1,AFFAIRE2,AFFAIRE3,AVENANT,TIERS, AFFAIRE0:THCritMaskEdit;TestCle:Boolean; Action : TActionFiche; TobAffaire,TobAffaires : Tob; bAvecMessage, bAvecAffModele, BProposition:boolean) : integer;
Function  IdentifiantAffaireRenseigne (Part1,Part2,Part3 : string): Boolean;
Procedure RAZIdentifiantAffaire (Var Part1,Part2,Part3 : string);
Function Existeaffaire (CodeAffaire,CodeTiers : String) : Boolean;
Function ExisteTiers (CodeTiers : String) : Boolean;
Function RechercheAffaire (Aff0,Aff1,Aff2,Aff3,Aff4 : String; EnCreatAff : Boolean):Boolean;
// Gestion des paramètres du code affaire
Procedure InitParamCodeAffaire ;
Procedure ModifParamCodeAffaire ( TypeExercice : String) ;
Procedure InitParpieceAffaire;
Procedure InitUneParpieceAffaire(TypePP : string);
Procedure ModifUneParPiece(TypePP : string);
Procedure ModifLesParpieceAffaire;
Procedure InitSoucheAffaire;
function PropositionToAffaire(CodeProposition:string; var Aff0, Aff1, Aff2, Aff3, Aff4:string; bGarderCpt,bUpdateCpt : Boolean):string;
Function AFRemplirTOBAffaire ( CodeAffaire : String ; TOBAffaire, TOBAffaires : TOB ) : boolean ;

// Fonctions sur Echéances affaires et total affaire
function  EvaluationNbEcheances(CodeAff, TypeCalcul : string; zinterval: integer; DateDebut , DateFin : TdateTime ):integer;
Procedure CumulEcheances(CodeAff:string;var zcum:double;var zpou:double;var znbr:integer);
Procedure ChargeEcheances(CodeAff,CodeGenerauto,typech:string;DateDebut,DateFin:TdateTime;TobEch:Tob; AvecFactures:boolean=false);
Function  RemplirTOBEcheance ( CodeAffaire : String;numech : Integer;TobEch : TOB ):boolean;
Function ExisteEcheanceAffaire ( CodeAffaire : String; facture : string='') : Boolean;
Function ExisteEcheanceBudget ( CodeAffaire : String) : Boolean;
Procedure ConvertDevToPivot(DEV : RDEVISE; MtDev : Double;  var MtPivot : double);
Procedure ConvertPivotToDev(DEV : RDEVISE; MtPivot : Double;  var  MtDev : double);
Function  CalculTotalAffaire(CodeAff, TypeGener, TypePrevu : String; TotalHTLigne : double; TobEcheances: TOB; var NbEche : integer; var baseEche : Boolean; DateDebGener :TDatetime ) : Double;
Function  EcheanceMoisAffaire (CodeAff : String; DateMois : TDatetime) : Double;
Procedure InverseEcheanceContre (CodeAff : String ; SaisieContre : Boolean);
function  CreerFactAff (TobAffaire : TOB; DateFact : TDateTime ;ztyp : string ): Boolean;

// Fonctions de liaison Affaire Piece
Function  CreerPieceAffaire(CodeTiers, CodeAffaire, StatutAffaire, Etablissement : String ; Var CleDocAffaire : R_CLEDOC; EnHT,SaisieContre : Boolean ):Boolean;
Function  DetruitPieceAffaire(CodeTiers, CodeAffaire :String ; CleDocAffaire : R_CLEDOC):Boolean;
Function  SelectPieceAffaire(CodeAffaire, StatutAffaire : String; Var CleDocAffaire : R_CLEDOC; Accepte : Boolean=False):Integer;
Function  GetTobPieceAffaire (CodeAff ,StatutAffaire : String; TobPiece : TOB) : Boolean;
Function  MajCodeAffaireSurPiece (OldAff,NewAff, NatPiece,NumPiece,Souche : string):integer;
Procedure ChargeTreeViewPieceAffaire ( TobPieceAffaire : TOB ; TVPiece:TTreeView ; CodeAffaire: String ) ;
Function  RecupPRPieceAffaire ( TobPiece : TOB) : Double;
Procedure AccesEntetePieceAffaire(TobPiece,TobAdresses : TOB ; Action : TActionFiche);
Function  SelectPieceAffaireBis(CodeAffaire, StatutAffaire: String; Var CleDocAffaire : R_CLEDOC ;var Sttiers,StRep:string ):Integer;

// Marge sur affaire
Function CalculMargesurAffaire (TotalHT, TotalPR : Double; QualMarge : string) : Double;
// Fonctions recup sur affaire depuis la piece
Procedure LectureLignesAffaire (Tobpiece : Tob; AvecEntete : Boolean);


// Fonctions appelées lors de la validation de piece
Procedure MajGestionAffaire (TobPiece , TobPiece_ori,TOBPieceRG,TOBBasesRG,TOBAcomptes: Tob;typeaction : string; DEV : Rdevise; FinTravaux : boolean=false);
Procedure MajAffaire(TobPiece,TOBAcomptes : Tob; CodeAffaireAvenant, typeaction : string;Action : TActionFiche; Duplic:Boolean;SaisieAvanc:boolean; CodeEtat : String='');
Procedure MajAffairePiece (Tobpiece ,TobPiece_ori: Tob;typeaction : string);
Procedure MajActivite (Tobpiece , TobPiece_ori: Tob;typeaction : string);
Procedure MajEcheance (Tobpiece, TobPiece_ori : Tob;typeaction : string);
Procedure MajFactEclat (Tobpiece : Tob);
Procedure ChargeTobActivite (var TobProd:TOB;TobPiece : TOB;Aff:array of string;Nbaff:integer );
Procedure EcrireDansAfCumul (TobProd,TobPiece : TOB ;Aff:array of string;Nbaff:integer);
procedure AjoutTypeArticle(racine:string;var Zwhere:string);

// Fonction de calcul des dates d'exercices comptables sur les missions
function CalculDateExercice (TypeExer : string; MoisCloture: integer; Exercice: String; Var DebutExer,Finexer,DebutInter,FinInter,DebutFact,FinFact,Liquidative,Garantie,Cloture : TDateTime; SurExercice : Boolean) : Boolean;
function CalculDateAffaire (DebutInter,FinInter : TDatetime; Var DebutFact,FinFact,Liquidative,Garantie,Cloture : TDateTime) : Boolean;

Function InterpreteZoneExercice (TypeExer : string; var Exercice : String; Var AnneeDeb,AnneeFin,Mois : integer; Obligatoire : Boolean) : Boolean;
function FormatZoneExercice (TypeExer : string; bTiretStocke : Boolean = False) : String;

Function GetChampsAffaire (codeAff, NomChamp : String) : string;
function ControleAffaire(CodeAff, TitleMsg,TypeSaisie : String): Boolean;
Function ValidationCodeEtatChantier(TypeSaisie, CodeEtat, TitleMsg : String) : Boolean;
function ValidationCodeEtatAppel(TypeSaisie, CodeEtat, TitleMsg : String) : boolean;
Function ValidationCodeEtatContrat(TypeSaisie, CodeEtat, TitleMsg : String): boolean;


// Tiers spécif Gestion interne
Function GetMoisClotureTiers (CodeTiers : String) : Integer;
// mcd 15/07/02 passage de GetChampsComplementTiers dans TiersUtil
Function IsTiersFerme (CodeTiers : String): Boolean;

Procedure AfficheTotauxDansLibelle (Sender: TObject; Ecran : TOBJECT );
Function FabriqueWhereNatureAuxiAff ( Statut : String) : String;

// Gestion des Avenants , etats sur affaires ...
function GetPlusEtatAffaire(bAffaire : Boolean) : String;
// Gestion des statuts
Function StatutReduitToComplet (stStatut : String) : String;
Function StatutCompletToReduit (stStatut : String) : String;

// Analyse de la zone ACTIVITEREPRISE
procedure Trt_Activiterepris(zrep : string; var zcombo : string;var toppre,topfou,topfra : boolean);

// fct pour blocage mono
Function ToutSeulAff () : Boolean;
Function BlocageAffaire( AcsEvtCourant, AcsAffaireCourante, AcsGroupeUser : string; AcdDateCourante : TDateTime; AcbBlocParlant,AcbAlerteParlant,AcbTous : boolean; var AvdDebutItv, AvdFinItv :TDateTime; AtobAffaires:TOB) : T_TypeBlocAff;
function TOBEtatsBloquesAffaire (AcsEvtCourant, AcsAlerteON, AcsGroupeUser : string) : TOB;
function ActionAffaireString( AcbModifTOMAutorisee:boolean;  TypeAction:TActionFiche ):string;

// fct générale pour GA
function RecupLibelleTiers(code : string): string;
Procedure ColCodeAffaireGrid (GS : THGRID;LesColonnes:string);
Function AFMajProspectClient (CodeTiers :string):Boolean;


Procedure RechAdresseFactAffaire ( TOBTiers,TOBAdresses,TOBPiece,TOBPieceAffaire  : TOB ) ;

// Lien d'une ligne piéce avec une tache
Procedure AFLignePieceToTache ( pStCodeAffaire,pStCodeTiers,pStLignePieceAff : String ; pAction:TActionFiche);
function IsTypeAffaire (Affaire : string; TypeAffaire : string) : boolean;

// OPTIMISATIONS LS
procedure ReinitTOBAffaires;
procedure StockeCetteAffaire (CodeAffaire : string); overload;
procedure StockeCetteAffaire (TOBAffaire : TOB); overload;
function FindCetteAffaire (CodeAffaire : string) : TOB;
function ControleDocumentsAffaire (CodeAffaire : string) : boolean;
function CalculMtEcheanceContrat (MtDoc: double ;CalcEcheance,stPeriode: string;Interval,NbIt : integer) : double;
function ConstitueCodeAffaireSql (Part0,Part1,Part2,Part3,Avenant : string) : string;

// --
//function GetEtatAppelsNonTermine : string;

implementation
uses utofAfPiece;

//**********************************************************************************
//**************** Fonctions de Traitement du Découpage du Code affaire ************
//**********************************************************************************

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 02/02/2000
Modifié le ... : 02/02/2000
Description .. : Gestion des composants des 3 parties du code affaire et découpage du code affaire concaténé
Mots clefs ... : AFFAIRE;CODE AFFAIRE
*****************************************************************}
function ChargeCleAffaire(Part0,Part1,Part2,Part3,Avenant: THCritMaskEDIT;BRechAffaire:TToolBarButton97; FTypeAction:TActionFiche; CodeAff:string;SaisieAffaire:Boolean) : integer;
var i,lng,NbPartieVisible, PositAvenant : integer;
    TypeCode : string3;
    Valeur,Libelle,CodePartie0,CodePartie1,CodePartie2,CodePartie3,CodeAvenant : String;
    Visible,Bproposition : Boolean;
    Code : THCritMaskEdit;
    OnChangetmp1, OnChangetmp2 : TNotifyEvent;
    HH            : THLabel ;
    FF            : TForm ;
    OK_SaisieAppel: Boolean;
    StPrefixe     : String;
Begin
  lng := 0;
  Visible :=False;
  result := 0;

  Ok_SaisieAppel := false;
  Ok_SaisieAppel := GetParamSocSecur('SO_SAISIECODEAPPEL', False);

  if part0 <> nil then StPrefixe := Part0.text;

  if Not(ctxAffaire in V_PGI.PGIContexte) and Not(ctxGCAFF in V_PGI.PGIContexte) then
    BEGIN
      if Part0  <> Nil then Part0.Visible := False;
      if Part1  <> Nil then Part1.Visible := False;
      if Part2  <> Nil then Part2.Visible := False;
      if Part3  <> Nil then Part3.Visible := False;
      if Avenant<> Nil then Avenant.Visible := False;
      if Part1<>Nil then
        BEGIN
          FF:=TForm(Part1.Owner) ;
          for i:=0 to FF.ComponentCount-1 do if FF.Components[i] is THLabel then
            BEGIN
            HH:=THLabel(FF.Components[i]) ;
            if HH.FocusControl=Part1 then
               BEGIN
               HH.Visible:=False ;
               Break ;
               END ;
            END ;
         END;
      Exit;
   END;

	if Part1 = nil then Exit;
//  if Part2 <> nil then Exit;

  OnChangetmp1 := Part1.OnChange;
	OnChangetmp2 := Part2.OnChange;

  Part1.OnChange := Nil;
  Part2.OnChange := Nil;

  Part1.Text :='' ;
  Part2.Text := '';
  Part3.Text := '';

  try
  NbPartieVisible := 0;
  PositAvenant := 0;
  // Gestion des affaires
  if Part0 <> Nil then
     BEGIN
     Part0.DataType:= 'AFSTATUTREDUIT';
     Part0.ElipsisButton:=True;
     CodePartie0 := Part0.text;
     END
  Else
	   CodePartie0 := 'A';

  if CodePartie0 = 'P' then bProposition := true else bProposition := False;

  for i:=1 to NbMaxPartieAffaire do
     Begin
     // recup des valeurs de la partie du code traitée
     case i of
          1: Begin
             Code:=Part1;
             if (StPrefixe = 'W') AND (not Ok_SaisieAppel) then
             		begin
                lng:=GetParamSoc('SO_APPCO1LNG');
                TypeCode:=GetParamSoc('SO_APPCO1TYPE');
                Valeur:=GetParamSoc('SO_APPCO1VALEUR');
           		  Visible:=True;
                Libelle:=GetParamSoc('SO_APPCO1LIB');
                end
             else
           	  	Begin
                lng:=VH_GC.CleAffaire.Co1Lng;
           		  TypeCode:=VH_GC.CleAffaire.Co1Type;
           	   	Valeur:=VH_GC.CleAffaire.Co1valeur;
           	  	Visible:=True;
           	  	Libelle:=VH_GC.CleAffaire.Co1Lib;
           	  	End;
             End;
          2: Begin
             Code:=Part2;
             if (StPrefixe = 'W') AND (not Ok_SaisieAppel) then
            		begin
                lng:=GetParamSoc('SO_APPCO2LNG');
                TypeCode:=GetParamSoc('SO_APPCO2TYPE');
                Valeur:=GetParamSoc('SO_APPCO2VALEUR');
           	   	Visible:=GetParamSoc('SO_APPCO2VISIBLE');
                Libelle:=GetParamSoc('SO_APPCO2LIB');
                end
             else
           		  Begin
                lng:= VH_GC.CleAffaire.Co2Lng;
                TypeCode:=VH_GC.CleAffaire.Co2Type;
                Valeur:=VH_GC.CleAffaire.Co2valeur;
                Visible:=VH_GC.CleAffaire.Co2Visible;
                Libelle:=VH_GC.CleAffaire.Co2Lib;
                End;
             End;
          3: Begin
	           Code:=Part3;
             if (StPrefixe = 'W') AND (not Ok_SaisieAppel) then
            		begin
                lng:=GetParamSoc('SO_APPCO3LNG');
                TypeCode:=GetParamSoc('SO_APPCO3TYPE');
                Valeur:=GetParamSoc('SO_APPCO3VALEUR');
             		Visible:=GetParamSoc('SO_APPCO3VISIBLE');
                Libelle:=GetParamSoc('SO_APPCO3LIB');
                end
             else
           	  	Begin
                lng:=VH_GC.CleAffaire.Co3Lng;
                TypeCode:=VH_GC.CleAffaire.Co3Type;
                Valeur:=VH_GC.CleAffaire.Co3valeur;
                Visible:=VH_GC.CleAffaire.Co3Visible;
                Libelle:=VH_GC.CleAffaire.Co3Lib;
                End;
             End;
          else
          code := nil;
      End;

      //FV1 - 27/06/2016 - FS#2039 - MATFOR - Modification possible du code affaire même si affaire acceptée
      If (i <= VH_GC.CleAffaire.NbPartie) Then
         Begin   // Partie utilisée
         Code.ReadOnly  := False;
         Code.enabled   := True;
         Code.color     := clWindow;
         Code.Visible   := Visible;
         Code.Hint      :=Libelle; // propriété visible + conseil
         Code.ElipsisButton:=False;
         Code.MaxLength :=lng;
         Code.Width     :=(lng * 8)+10; // Calcul longueur des zones
         if ((Typecode='SAI') And (i=2) And (VH_GC.AFFORMATEXER<>'AUC')) then
            Code.Width := Code.Width + 15; // séparateurs en plus
         If (i=2) Then Code.Left:=Part1.Left+Part1.Width+5;
         If (i=3) Then Code.Left:=Part2.Left+Part2.Width+5;
         If (SaisieAffaire) then
            Begin
              if (FTypeAction=taCreat) then
              begin
                 Code.ReadOnly:= False;
                 Code.enabled := true;
              end
              else
              begin
                 Code.ReadOnly:= True;
                 Code.enabled := False;
              end;
            End
         Else // Code affaire depuis une autre saisie
            Begin
              If (FTypeAction <> taConsult) Then
              begin
                 Code.ReadOnly:= False;
                 Code.enabled := True;
              end
              else
              begin
                 Code.ReadOnly:= True;
                 Code.enabled := False;
              end;
            End;
         // spécificitées en fonction du type de zone ...
         if (Typecode= 'FIX')then
            begin
            Code.Width:=45;
            Code.ReadOnly:= False;
            Code.enabled := True;
            if SaisieAffaire then
            Begin
            	Code.enabled := False;
            	Code.Text    := valeur;
            end
            else
            	begin
              Code.ReadOnly:= True;
            	Code.enabled := False;
              Code.Text    := '';
            	end;
            end
         Else if (Typecode= 'DAT')  then
          	Begin
            Code.EditMask := FormatZoneExercice(VH_GC.AFFormatExer,false);
            if length(valeur) = 6 then
  	  	       Code.Hint := Code.Hint+ ' Année-Mois (yyyymm)'
            else
						   Code.Hint := Code.Hint+ ' Année Fin-Mois (yymm)'
            end
         else if (Typecode= 'LIS') then                                       // Liste de saisie
            Begin
            if ((SaisieAffaire) And
               (FTypeAction = taCreat )) Or
               (Not(SaisieAffaire) And
               (FTypeAction <> taConsult) And
               (Code.ReadOnly<>True)) Then
               Code.ElipsisButton:=True;
            Code.DataType:='AFFAIREPART'+intToStr(i);
            Code.Width:=45;
            End
         else if ((pos(Typecode,'CPT;CPM')>0) And SaisieAffaire) then // compteur
             BEGIN
               if TypeCode = 'CPT' Then
               begin
                Code.ReadOnly   := True;
                Code.Enabled    := True;
                if (FTypeAction=taCreat) then
                  Code.color    := ClBtnFace
                else
                  Code.color    := ClWindow;
               end else
               begin // TYpeCode = CPM
                if (FTypeAction=taCreat) then
                begin
                	Code.ReadOnly := False;
                  Code.enabled  := True;
                  Code.color    := ClWindow;
                end else
                begin
                	Code.ReadOnly := True;
                  Code.enabled  := False;
                  Code.color    := ClBtnFace
                end;
               end;
             End
          else if ((Typecode='SAI') And (i=2) And (VH_GC.AFFormatExer<>'AUC')) then  // Formattage Exercice
             BEGIN
             Code.EditMask := FormatZoneExercice(VH_GC.AFFormatExer,false);
             if (VH_GC.AFFormatExer='AA') then
                 Code.Hint :=Code.Hint+ '(Année Début-Année fin d''exercice)'
             else if (VH_GC.AFFormatExer='A') then
                 Code.Hint :=Code.Hint+ '(Millésime : Année de fin d''exercice)'
             else
                 Code.Hint :=Code.Hint+ '(Année -Mois Fin d''exercice)';
             END;
              // mcd déplacer, car longueur change si LIS 26/06/01
          if Visible then
             BEGIN
             PositAvenant := Code.Left +  Code.Width + 5; // Calcul position avenant
             NbPartieVisible := i;
          	 Result := Code.Left + Code.Width + 5;
             //Result := Code.Left + Code.Width + 5;
             END;
        End
      Else
        Begin  // RAZ de la partie de structure non utilisée
          Code.Visible:=False;
        End;
    End;

  CodePartie1 := '';
  CodePartie2 := '';
  CodePartie3 :='';
  CodePartie0 := '';
  CodeAvenant := '';

  if (codeAff <> '')  then
     begin
     {$IFDEF BTP}
     BTPCodeAffaireDecoupe(CodeAff,CodePartie0,CodePartie1,CodePartie2,CodePartie3,CodeAvenant,FTypeAction,SaisieAffaire);
     {$ELSE}
     CodeAffaireDecoupe(CodeAff,CodePartie0,CodePartie1,CodePartie2,CodePartie3,CodeAvenant,FTypeAction,SaisieAffaire);
     {$ENDIF}
     Part1.Text := CodePartie1 ;
     Part2.Text := CodePartie2;
     Part3.Text := CodePartie3;
     end;

  if (Part0<>nil) then Part0.Text :=CodePartie0 ;
  // mcd 18/09/01 dépalcer car non OK si partie FIX Part1.Text :=CodePartie1 ; Part2.Text := CodePartie2; Part3.Text := CodePartie3;

  // positionnement du bouton de recherche affaire
  If ((NbPartieVisible > 0) And (BRechAffaire <> nil)) Then
    Begin
      Case  NbPartieVisible of
          1 : BRechAffaire.Left := Part1.Left + Part1.Width+5;
          2 : BRechAffaire.Left := Part2.Left + Part2.Width+5;
          3 : BRechAffaire.Left := Part3.Left + Part3.Width+5;
      End;
    End;

  // Gestion de la zone avenant
  if Avenant <> Nil then
    BEGIN
      if (VH_GC.CleAffaire.GestionAvenant) then
          BEGIN
          Avenant.Visible := true;
          Avenant.Left := PositAvenant;
          Avenant.Width :=2 * 10 + 5;
          Avenant.Hint := 'Avenant';
          Avenant.MaxLength := 2;
          {$IFDEF BTP}
          Avenant.Text := BTPTestcodeAvenant(CodeAvenant,SaisieAffaire);
          {$ELSE}
          Avenant.Text := TestcodeAvenant(CodeAvenant,SaisieAffaire);
          {$ENDIF}
        	Result := PositAvenant + Avenant.Width + 5;
          //Result := PositAvenant + Avenant.Width + 5;
          If (SaisieAffaire) then
              BEGIN
              Avenant.ReadOnly:=True;
              if (FTypeAction=taCreat) then
                 Avenant.color := ClBtnFace
              else
                 Avenant.color := ClWindow;
              END;
          if BRechAffaire <> nil then
          begin
          	BRechAffaire.Left := Avenant.Left + Avenant.Width+5;
          end;
          END
      else Avenant.Visible := false;
    END;

  finally
    Part1.OnChange:= OnChangetmp1;
    Part2.OnChange:= OnChangetmp2;
  end;
End;

Procedure CalculeCodeAffaire(Part0,Part1,Part2,Part3,Avenant: THCritMaskEDIT;BRechAffaire:TToolBarButton97; FTypeAction:TActionFiche; CodeAff:string;SaisieAffaire:Boolean;FromAppel:boolean=false);
var   OK_SaisieAppel: Boolean;
    	Bproposition : Boolean;
    	CodePartie0,CodePartie1,CodePartie2,CodePartie3,CodeAvenant : String;
    	Code : THCritMaskEdit;
      i : integer;
    	TypeCode : string3;
begin
  Ok_SaisieAppel := false;
  Ok_SaisieAppel := GetParamSocSecur('SO_SAISIECODEAPPEL', False);
//
	if Part1 = nil then Exit;
  if Part0 = Nil then CodePartie0 := 'A' else CodePartie0 := Part0.text;
  if CodePartie0 = 'P' then bProposition := true else bProposition := False;

  for i:=1 to NbMaxPartieAffaire do
  Begin
  	If (i > VH_GC.CleAffaire.NbPartie) Then exit;
     // recup des valeurs de la partie du code traitée
		case i of
      1: Begin
         Code:=Part1;
         TypeCode:=GetParamSoc('SO_AFFCO1TYPE');
         if TypeCode = 'CPM' then
         begin
          Code.Text := CompteurPartieAffaire(i, bProposition,true,FromAppel);
         end;
       end;
      2: Begin
         Code:=Part2;
         TypeCode:=GetParamSoc('SO_AFFCO2TYPE');
         if TypeCode = 'CPM' then
         begin
					Code.Text := CompteurPartieAffaire(i, bProposition,True,FromAppel);
         end;
       end;
      3: Begin
         Code:=Part3;
         TypeCode:=GetParamSoc('SO_AFFCO3TYPE');
         if TypeCode = 'CPM' then
         begin
          Code.Text := CompteurPartieAffaire(i, bProposition,True,FromAppel);
         end;
			end;
		end;
  end;
end;


Function MaxLngAffaireAffiche : integer;
BEGIN
Result := VH_GC.CleAffaire.Co1Lng;
if (VH_GC.CleAffaire.Co2Visible) And (VH_GC.CleAffaire.NbPartie>1) then
   begin
   Result := Result + VH_GC.CleAffaire.Co2Lng;
          // mcd 18/09/01 séparateur même en GA if ctxscot in V_PGI.PGIContexte then Inc(Result);
   Inc(Result);
        //mcd 30/09/03 si Format exercice AA ou AM, il faut en plus ajouter le - du ocde exercie
   If (GetParamsoc('SO_AFFORMATEXER')='AA') or (GetParamsoc('SO_AFFORMATEXER')='AM') then Inc(Result);
   end;
if (VH_GC.CleAffaire.Co3Visible) And (VH_GC.CleAffaire.NbPartie>2) then
   begin
   Result := Result + VH_GC.CleAffaire.Co3Lng;
   Inc(Result);
   end;
END;

Function CodeAffAfficheToComplet (CodeAffReduit,Tiers : string) : string;
Var BesoinSQL : Boolean;
    Part0,Part1,Part2,Part3,Avenant,stWhere : String;
    Q : TQuery;
BEGIN
Result := '';
if CodeAffReduit='' then Exit;
BesoinSQL := False;
// requête nécessaire si une partie n'est pas visible par l'utilisateur
if Not(VH_GC.CleAffaire.Co2Visible) And (VH_GC.CleAffaire.NbPartie>1) then BesoinSQL := True;
if Not(VH_GC.CleAffaire.Co3Visible) And (VH_GC.CleAffaire.NbPartie>2) then BesoinSQL := True;

Part0 := 'A';
Part1 := Copy(CodeAffReduit,1, VH_GC.CleAffaire.Co1Lng);
if (VH_GC.CleAffaire.Co2Visible) And (VH_GC.CleAffaire.NbPartie>1) then
   Part2 := Copy(CodeAffReduit,VH_GC.CleAffaire.Co1Lng+1,VH_GC.CleAffaire.Co2Lng);
if (VH_GC.CleAffaire.Co3Visible) And (VH_GC.CleAffaire.NbPartie>2) then
   Part3 := Copy(CodeAffReduit,VH_GC.CleAffaire.Co1Lng+VH_GC.CleAffaire.Co2Lng+1,VH_GC.CleAffaire.Co3Lng);
Avenant := '00';

if (BesoinSQL) then
   BEGIN
   stWhere := '';
   if Part2 <> '' then stWhere := ' AND AFF_AFFAIRE2="'+ Part2 + '"';
   if Part3 <> '' then stWhere := ' AND AFF_AFFAIRE3="'+ Part3 + '"';
   if Tiers <> '' then stWhere := stWhere + ' AND AFF_TIERS="'+Tiers+ '"';
// Modif PL le 09/01/02 pour optimiser le nombre d'occurrences ramenées en e-agl
   Q:= OpenSQL ('SELECT AFF_AFFAIRE2,AFF_AFFAIRE3 FROM AFFAIRE WHERE AFF_AFFAIRE1="'+ Part1+ '"' + stWhere,True,2,'',true);
//   Q:= OpenSQL ('SELECT AFF_AFFAIRE2,AFF_AFFAIRE3 FROM AFFAIRE WHERE AFF_AFFAIRE1="'+ Part1+ '"' + stWhere,True);
   try
   if Not(Q.EOF) then
      BEGIN
      if Not(VH_GC.CleAffaire.Co2Visible) And (VH_GC.CleAffaire.NbPartie>1) then
         Part2 := Q.FindField('AFF_AFFAIRE2').AsString;
      if Not(VH_GC.CleAffaire.Co3Visible) And (VH_GC.CleAffaire.NbPartie>2) then
         Part3 := Q.FindField('AFF_AFFAIRE3').AsString;
      // si plus d'une affaire raz du code
      Q.Next; if (Q.EOF) then begin ferme(Q); Exit; end;// pas d'affaire ou plus d'une affaire correspondante
      END;
   finally
   Ferme(Q);
   end;
   END;
Result := RegroupePartiesAffaire (Part0, Part1, Part2, Part3,Avenant );
END;

{***********A.G.L.***********************************************
Auteur  ...... : PATRICE ARANEGA
Créé le ...... : 21/05/2000
Modifié le ... : 21/05/2000
Description .. : Retourne code affaire sans 1er car et avenant
Mots clefs ... : AFFAIRE;CODE AFFAIRE
*****************************************************************}
Function CodeAffaireUser (CodeAffaire : String) : String;
BEGIN
Result := Copy(CodeAffaire,2,14);
END;

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 02/02/2000
Modifié le ... : 02/02/2000
Description .. : Contrôle des 3 parties du code affaire et composition du code affaire concaténé.
Mots clefs ... : AFFAIRE;CODE AFFAIRE
*****************************************************************}
function DechargeCleAffaire(Part0,Part1,Part2,Part3,Avenant: THCritMaskEDIT; CodeTiers : string;
                              FTypeAction:TActionFiche; SaisieAffaire,AfficheMsg,bProposition:Boolean;
                              var AviPartErreur:integer):String; // Modif PL le 15/06/2000 : ajout de AviPartErreur pour savoir quelle zone sort en erreur
Var i               : Integer;
    AnneeDeb        : Integer;
    AnneeFin        : Integer;
    Mois            : integer;
    Code            : THCritMaskEdit;
    TypeCode        : String3;
    Exercice        : String;
    ExerObligatoire : Boolean;
    FromAppel       : Boolean;
    StatutReduit    : string;
    CodeAvenant     : string;
    P1, P2, P3      : string;
Begin
Result:='';
AviPartErreur:=0;

// Gestion avenant + statut
if (Avenant <> Nil) then
   {$IFDEF BTP}
   CodeAvenant := BTPTestCodeAvenant (Avenant.Text,True)
   {$ELSE}
   CodeAvenant := TestCodeAvenant (Avenant.Text,True)
   {$ENDIF}
else
   CodeAvenant := '00';
if Part0 <> Nil then
    StatutReduit := Part0.Text
else
    BEGIN
    if bProposition then StatutReduit := 'P' else StatutReduit := 'A';
    END;

// En création de mission test qu'une même mission / exercice n'existe pas pour ce client
 // mcd 18/09/01 if (ctxScot in V_PGI.PGIContexte) And (FTypeAction=taCreat) And (SaisieAffaire) then
if (GetParamSoc('SO_AFFORMATEXER') <>'AUC') And (FTypeAction=taCreat) And (SaisieAffaire) then
    BEGIN
    if ChercheMissionExercice (StatutReduit,Part1.Text,Part2.Text, CodeAvenant,CodeTiers) then
        BEGIN
        If AfficheMsg Then PGIBoxAF('Code affaire déja utilisé pour ce client',TitreHalley);
        Exit;
        END;
    END;
// En création d'affaires test qu'une affaire n'existe sur ce code
   // mcd 18/09/01 if Not(ctxScot in V_PGI.PGIContexte) And (FTypeAction=taCreat) And (SaisieAffaire) then
if (GetParamSoc('SO_AFFORMATEXER') ='AUC') And (FTypeAction=taCreat) And (SaisieAffaire) then
    BEGIN
    if RechercheAffaire (StatutReduit,Part1.Text,Part2.Text,Part3.Text, CodeAvenant,True) then
        BEGIN
        If AfficheMsg Then PGIBoxAF('Code affaire déja utilisé',TitreHalley);
        Exit;
        END;
    END;

for i:=1 to VH_GC.CleAffaire.NbPartie do
    Begin
    case i of
         1: Begin Code:=Part1; TypeCode:=VH_GC.CleAffaire.Co1Type; End;
         2: Begin Code:=Part2; TypeCode:=VH_GC.CleAffaire.Co2Type; End;
         3: Begin Code:=Part3; TypeCode:=VH_GC.CleAffaire.Co3Type; End;
    else Code := nil;
    End;
    //
    if ((FTypeAction=taCreat) And (SaisieAffaire)) then
    Begin
      //FV1 : 10/11/2015 - FS#1781 - VEODIS : N° d'affaire et N° d'appel ne sont plus sur le même Compteur
      if Part0.text = 'W' then FromAppel   := not GetParamsocSecur('SO_CPTFORAFF', false);
      //
      if (Pos(TypeCode, 'CPT;CPM')>0) then
      begin
        if Code.text = '' then Code.Text := CompteurPartieAffaire(i, bProposition,True,FromAppel);
      end;
      if (Code.Text ='') then
      Begin      // Si une partie n'est pas renseignée code invalide...
        If AfficheMsg Then PGIBoxAF('Partie du code affaire non renseignée',TitreHalley);
        if Code.CanFocus then   Code.SetFocus;
        Exit;
      End;
    End;

    if ((TypeCode = 'LIS') and (Code.Text<>'') and (StatutReduit<>'W')) then    // Contrôle la cohérence du code / la liste
        BEGIN
        if (Not LookUpValueExist(Code)) then
            Begin
            Code.Text:='';
            if AfficheMsg then
               BEGIN
               PGIBoxAF('Valeur incorrecte',TitreHalley);
            if Code.CanFocus then   Code.SetFocus;
            AviPartErreur:= i;
            Exit;
            END;
            End;
        END;

    if ((Typecode='SAI') and (i=2) and (VH_GC.AFFORMATEXER<>'AUC')) then
        BEGIN
        Exercice := Code.Text;
        ExerObligatoire := False;
        if (FTypeAction=taCreat) And (SaisieAffaire) then ExerObligatoire := True;
        if Not (InterpreteZoneExercice(VH_GC.AFFORMATEXER,Exercice,AnneeDeb,AnneeFin,Mois,ExerObligatoire)) then
            BEGIN
            if AfficheMsg then BEGIN PGIBOXAF('Exercice incorrect ',TitreHalley);
            if Code.CanFocus then   Code.SetFocus;
            AviPartErreur:= i;
            Exit;
            END;
        END
        else Code.Text:=Exercice;
        END;
    End;
// recup des valeurs des parties du code traité
P1:= Part1.Text;
P2:= Part2.Text;
P3:= Part3.Text;
Result := RegroupePartiesAffaire (StatutReduit, P1, P2, P3, CodeAvenant);
if (Avenant <> Nil) then Avenant.Text := CodeAvenant;
End;



{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 17/02/2000
Modifié le ... : 17/02/2000
Description .. : Regroupement des 3 parties du code affaire à partir de string + test validité
Mots clefs ... : AFFAIRE;CODE AFFAIRE
*****************************************************************}
function  CodeAffaireRegroupe(Var Part0, Part1 ,Part2 ,Part3, Avenant : string ; FTypeAction:TActionFiche ; SaisieAffaire,bGarderCpt,bUpdateCpt : Boolean): String;
Var codepartie    : String;
    SQL           : String;
    TypeTab       : String;
    i             : integer;
    TypeCode      : String3;
    bProposition  : Boolean;
    FromAppel     : Boolean;
Begin
  Result:='';

  for i:=1 to VH_GC.CleAffaire.NbPartie do
  Begin
    case i of
         1: Begin CodePartie:=Part1; TypeCode:=VH_GC.CleAffaire.Co1Type; TypeTab := 'AP1'; End;
         2: Begin CodePartie:=Part2; TypeCode:=VH_GC.CleAffaire.Co2Type; TypeTab := 'AP2'; End;
         3: Begin CodePartie:=Part3; TypeCode:=VH_GC.CleAffaire.Co3Type; TypeTab := 'AP3'; End;
    End;
    Trim(codePartie);
    FromAppel    := False;
    //
    if ((FTypeAction=taCreat) And (SaisieAffaire)) then
    Begin
      if Part0 = 'P' then bProposition := true else bProposition := False;
      //FV1 : 16/06/20015 - FS#1625 - SOTRELEC - Numéro de compteur code appel fonctionne plus en automatique
      //FV1 : 10/11/2015 - FS#1781 - VEODIS : N° d'affaire et N° d'appel ne sont plus sur le même Compteur
      if part0 = 'W' then FromAppel := Not GetParamSocSecur('SO_CPTFORAFF', false);
      //
      //FV1 : 10/11/2015 - FS#1781 - VEODIS : N° d'affaire et N° d'appel ne sont plus sur le même Compteur
      //if (Pos(TypeCode ,'CPT;CPM')>0) then
      if (Pos(TypeCode, 'CPM')>0) then
      begin
        if Not(bGarderCpt) or (CodePartie = '') then
        begin
          //FV1 : 16/06/20015 - FS#1625 - SOTRELEC - Numéro de compteur code appel fonctionne plus en automatique
          CodePartie := CompteurPartieAffaire(i,bProposition,bUpdateCpt,FromAppel)
        end;
        if (CodePartie ='') then Exit; // Si une partie n'est pas renseignée code invalide...
      End;
      //
      if ((TypeCode = 'LIS') And (CodePartie<>'')) then    // Contrôle la cohérence du code / la liste
      Begin
        SQL:='SELECT CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="'+TypeTab+ '" AND CC_CODE="'+CodePartie+'"' ;
        if (Not ExisteSQL(SQL)) then Begin CodePartie:=''; End;
      End;
      //
      case i of
          1:  Part1 := CodePartie;
          2:  Part2 := CodePartie;
          3:  Part3 := CodePartie;
      End;
    End;
    // recup des valeurs des parties du code traité
    Result := RegroupePartiesAffaire (Part0, Part1, Part2, Part3,Avenant );
  End;

end;

Function CompteurPartieAffaire ( NumPartie : Integer; bProposition,bUpdateCpt : boolean; FromAppel:boolean=false  ) : string;
Var NewValeur,Valeur,suffixe : String;
    Cpt,lng : integer;
    Q : TQuery;
    StSQL : String;
Begin

  Result := ''; NewValeur := ''; Valeur := '';
  if (bProposition) And (VH_GC.CleAffaire.ProDifferent) then Suffixe := 'PRO' else Suffixe := '';
  case NumPartie of
      1: lng := VH_GC.CleAffaire.Co1Lng;
      2: lng := VH_GC.CleAffaire.Co2Lng;
      3: lng := VH_GC.CleAffaire.Co3Lng;
      else lng := 0;
      End;

  Q := nil;
  try
    if FromAppel then
    begin
      StSQL :='SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_APPCO' + IntToStr(NumPartie) + 'VALEUR'+ suffixe + '"';
      Q := OpenSQL(StSQL,True);
    end else
    begin
      StSQL := 'SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_AFFCO' + IntToStr(NumPartie) + 'VALEUR'+ suffixe + '"';
    	Q := OpenSQL(StSQL,True,-1,'',true);
    end;
    if Not Q.EOF then Valeur:= Q.Fields[0].asstring;

    // sortie rapide si une erreur est rencontrée
    if Valeur = '' then begin Result := ''; exit end;

    if length(Valeur) < lng then
        BEGIN
        Cpt := StrToInt(Valeur);
        Valeur := Format('%*.*d',[lng,lng,Cpt]);
        END;
    Result:= Valeur;

    NewValeur := CodeAffaireSuivant(Valeur, lng);
           //mcd 03/08/00 cas ou compteur faux ou depasse X éléments en démo
    if (NewValeur = '') then begin
       result:='';
       exit;
       end;
    // MAJ du compteur incrémenté en mémoire + Table
    if bUpdateCpt then
       begin
       if FromAppel then
       begin
          SetParamsoc('SO_APPCO'+IntToStr(NumPartie)+'VALEUR'+Suffixe, NewValeur );
          ExecuteSQL('UPDATE PARAMSOC SET SOC_DATA="'+NewValeur+'" WHERE SOC_NOM="SO_APPCO' + IntToStr(NumPartie) + 'VALEUR'+ suffixe + '"');
       end else
       begin
          SetParamsoc('SO_AFFCO'+IntToStr(NumPartie)+'VALEUR'+Suffixe, NewValeur );
          ExecuteSQL('UPDATE PARAMSOC SET SOC_DATA="'+NewValeur+'" WHERE SOC_NOM="SO_AFFCO' + IntToStr(NumPartie) + 'VALEUR'+ suffixe + '"');
       end;
       end
    else
       begin
    {$IFDEF DEBUG}
       PGIBoxAF('Attention valeur cpt affaire non enregistré n°: ' + NewValeur,'Debug Compteur affaire');
    {$ENDIF}
       end;

    if bProposition then
        BEGIN
        case NumPartie of
        1:  VH_GC.CleAffaire.Co1valeurPro := NewValeur ;
        2:  VH_GC.CleAffaire.Co2valeurPro := NewValeur ;
        3:  VH_GC.CleAffaire.Co3valeurPro := NewValeur ;
        End;
        END
    else if FromAppel then
         BEGIN
         case NumPartie of
            1:  SetParamSoc('SO_APPCO1VALEUR', NewValeur) ;
            2:  SetParamSoc('SO_APPCO2VALEUR', NewValeur) ;
            3:  SetParamSoc('SO_APPCO3VALEUR', NewValeur) ;
         End;
        END
     else BEGIN
        case NumPartie of
            1:  VH_GC.CleAffaire.Co1valeur := NewValeur ;
            2:  VH_GC.CleAffaire.Co2valeur := NewValeur ;
            3:  VH_GC.CleAffaire.Co3valeur := NewValeur ;
            End;
        END;
  finally
    Ferme(Q);
  end;
End;

function CodeAffaireSuivant ( Code : String; Lng : Integer ) : String ;
Var i ,cpt: integer ;
{$IFDEF BTP}
    Qr : TQuery;
    nbr:integer;
{$ENDIF}
BEGIN
if code = '' then code := '0';
Code:=AnsiUppercase(Trim(Code)) ;
if Copy(code,1,lng) < Copy('99999999999999999',1,lng) then
    BEGIN
        // code numerique, on incrémente
    Cpt := StrToInt(code);
    Inc(Cpt);
       //mcd 03/08/00
    if V_PGI.VersionDemo then begin
{$IFDEF BTP}
       Qr:=OpenSQL('SELECT COUNT(*) FROM AFFAIRE WHERE AFF_AFFAIRE0="A"',True,-1,'',true) ;
       if Not Qr.EOF then Nbr:=Qr.Fields[0].AsInteger else Nbr:=0;
       Ferme(Qr) ;
       if (Nbr > 50) then  begin
{$ELSE}
       if (cpt > 50) then  begin
{$ENDIF}
          PGIBoxAf('Version démonstration,Seulement 50 fiches affaires permises',TitreHalley);
          Result:='' ;
          exit;
          end;
       end;
    Code := Format('%*.*d',[lng,lng,Cpt]);
    Result:=Code ;
    END
else
    BEGIN
        // numérotation alpha des compteurs
    if Copy(code,1,lng) < Copy('ZZZZZZZZZZZZZZZZZ',1,lng) then
     begin
       if Length(Code) < Lng then Result:=Copy('AAAAAAAAAAAAAAAAA',1,Lng) else
         BEGIN
         i:=Lng ; While Code[i]='Z' do BEGIN Code[i]:='A' ; Dec(i) ; END ;
         Code[i]:=Succ(Code[i]) ;
         if Code[i]= ';' then Code[i]:=Succ(Code[i]) ;  // ; pose pb ds traitements !!!
         Result:=Code ;
         END;
       end
     else begin
          // Si arrivé au maxi, on considère que l'on peut repartir à 0
          // et que les missions ayant ces valeurs de compteurs seront supprimées
          // attention si compteur sur 3c ...
         Result:=Copy('00000000000000',1,lng);
         end;
    END;

END;

function ConstitueCodeAffaireSql (Part0,Part1,Part2,Part3,Avenant : string) : string;
var V0,V1,V2,V3,Valeur : string;
		II : Integer;
begin
  Valeur := '';

  {$IFDEF BTP}
  Avenant := BTPTestCodeAvenant(Avenant,True);
  Part0 := BTPTestStatutReduit(Part0); // 1er car. correspondant au statut
  {$ELSE}
  Avenant := TestCodeAvenant(Avenant,True);
  Part0 := TestStatutReduit(Part0); // 1er car. correspondant au statut
  {$ENDIF}

  V0 := Part0;
  If (VH_GC.CleAffaire.NbPartie >= 1) then
  begin
    if Part1 <> '' then
    begin
  		V1:=Format('%-*.*s',[VH_GC.CleAffaire.Co1Lng,VH_GC.CleAffaire.Co1Lng,Part1]);
    end else V1 := Copy(StringOfChar('_',30),1,VH_GC.CleAffaire.Co1Lng);
  end;
  If (VH_GC.CleAffaire.NbPartie >= 2) then
  begin
    if Part2 <> '' then
    begin
  		V2:=Format('%-*.*s',[VH_GC.CleAffaire.Co2Lng,VH_GC.CleAffaire.Co2Lng,Part2]);
    end else V2 := Copy(StringOfChar('_',30),1,VH_GC.CleAffaire.Co2Lng);
  end;
  If (VH_GC.CleAffaire.NbPartie >= 3) then
  begin
    if Part3 <> '' then
    begin
  		V3:=Format('%-*.*s',[VH_GC.CleAffaire.Co3Lng,VH_GC.CleAffaire.Co3Lng,Part3]);
    end else V3 := Copy(StringOfChar('_',30),1,VH_GC.CleAffaire.Co3Lng);
  end;
  Valeur := V0+V1+V2+V3;
  if GetParamSoc('SO_AFFGestionAvenant') then
  begin
  	Valeur := Format ( '%-*.*s',[15,15,Valeur]) + Avenant; // 2 car d'avenant à la fin du code affaire
  end else valeur := Valeur + '%';
  Result:=TrimRight(Valeur);

end;

Function RegroupePartiesAffaire (Var Part0, Part1, Part2, Part3,Avenant : String) : string;
Var Valeur : string;
BEGIN
Valeur := '';

{$IFDEF BTP}
Avenant := BTPTestCodeAvenant(Avenant,True);
Part0 := BTPTestStatutReduit(Part0); // 1er car. correspondant au statut
{$ELSE}
Avenant := TestCodeAvenant(Avenant,True);
Part0 := TestStatutReduit(Part0); // 1er car. correspondant au statut
{$ENDIF}

Valeur := Part0;
If (VH_GC.CleAffaire.NbPartie >= 1) then
    Valeur:=Valeur + Format('%-*.*s',[VH_GC.CleAffaire.Co1Lng,VH_GC.CleAffaire.Co1Lng,Part1]);
If (VH_GC.CleAffaire.NbPartie >= 2) then
    Valeur:=Valeur+Format('%-*.*s',[VH_GC.CleAffaire.Co2Lng,VH_GC.CleAffaire.Co2Lng,Part2]);
If (VH_GC.CleAffaire.NbPartie >= 3) then
    Valeur:=Valeur+Format('%-*.*s',[VH_GC.CleAffaire.Co3Lng,VH_GC.CleAffaire.Co3Lng,Part3]);
Valeur := Format ( '%-*.*s',[15,15,Valeur]) + Avenant; // 2 car d'avenant à la fin du code affaire
Result:=TrimRight(Valeur);
END;

function ChercheMissionExercice (StatutReduit, Partie1 , Partie2, Avenant, CodeTiers : string): Boolean;
Var Q : TQuery ;
    stWhere : string;
BEGIN
// On peut créer une mission sur le même code avec un autre code avenant
// On ne peut pas créer de mission sur le même code qu'une proposition ( phase d'acceptation uniquement)
stWhere := 'AFF_AFFAIRE1="'+ Partie1+ '" ';
if StatutReduit <> '' then
   stWhere := stWhere + 'AND AFF_AFFAIRE0="'+ StatutReduit+ '" ';
if (Pos(VH_GC.CleAffaire.Co2Type, 'CPT;CPM')=0) then
   stWhere := stWhere + 'AND AFF_AFFAIRE2="'+ Partie2+ '" ';

stWhere := stWhere + 'AND  AFF_TIERS="' + CodeTiers + '" AND AFF_AVENANT="'+ Avenant + '"';
// Modif PL le 09/01/02 pour optimisation du nb d'enregistrements lus en eagl
Q:= OpenSQL ('SELECT AFF_AFFAIRE FROM AFFAIRE WHERE ' + stWhere , True, 1,'',true) ;
//Q:= OpenSQL ('SELECT AFF_AFFAIRE FROM AFFAIRE WHERE ' + stWhere , True) ;
If (Q.EOF) then Result := False else Result := True;
Ferme(Q);
END;

function LookUpPartiesAffaire   ( NumPartie : Integer ; GS : THGrid ) : boolean;
Var     titre, typepart, typetablette : string;
Begin
Result := False;

case NumPartie of
    0: Begin titre := 'Statuts'; typepart := 'ASA'; typetablette:='CO'; End;
    1: Begin titre := VH_GC.CleAffaire.Co1Lib; typepart := 'AP1'; typetablette:='CC'; End;
    2: Begin titre := VH_GC.CleAffaire.Co2Lib; typepart := 'AP2'; typetablette:='CC'; End;
    3: Begin titre := VH_GC.CleAffaire.Co3Lib; typepart := 'AP3'; typetablette:='CC'; End;
    End;

if (typetablette='CC') then
   Result := LookupList(GS,titre, 'CHOIXCOD' ,'CC_CODE' ,'CC_LIBELLE' ,'CC_TYPE="'+ typepart + '"','CC_CODE',False,-1)
else
if (typetablette='CO') then
   Result := LookupList(GS,titre, 'COMMUN' ,'CO_ABREGE' ,'CO_LIBELLE' ,'CO_TYPE="'+ typepart + '"','CO_ABREGE',False,-1)
   ;
End;

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 02/02/2000
Modifié le ... : 02/02/2000
Description .. : Renvoi le nombre d'affaires existantes en fonction de la clé passée.
Mots clefs ... : AFFAIRE;CODE AFFAIRE
*****************************************************************}
Function TestCleAffaire(AFFAIRE,AFFAIRE1,AFFAIRE2,AFFAIRE3,AVENANT,TIERS:THCritMaskEdit; Part0 : string; bAvecMessage, bAvecAffModele,bProposition:boolean):integer;
Var stAffaire,stAffaire0,stAffaire1,stAffaire2,stAffaire3,stAvenant,stTiers : String;
Begin
stAffaire := AFFAIRE.Text;
stAffaire1:= AFFAIRE1.Text;
stAffaire2:= AFFAIRE2.Text;
stAffaire3:= AFFAIRE3.Text;
stAvenant := AVENANT.Text;
stTiers   := TIERS.Text;
stAffaire0:= Part0;
Result:= TeststCleAffaire(stAffaire,stAffaire0,stAffaire1,stAffaire2,stAffaire3,stAvenant,stTiers,bAvecMessage, bAvecAffModele,bProposition,true);
AFFAIRE.Text := stAffaire;
AFFAIRE1.Text:= stAffaire1;
AFFAIRE2.Text:= stAffaire2;
AFFAIRE3.Text:= stAffaire3;
AVENANT.Text := stAvenant;
TIERS.Text   := stTiers;
End;

Function TeststCleAffaire(var Affaire,Affaire0,Affaire1,Affaire2,Affaire3,Avenant,Tiers : string; bAvecMessage, bAvecAffModele,bProposition,WhereParties:boolean):integer;
Var Q : TQuery ;
    CodeAffaire,sReq,StatutReduit : String;
Begin

  Result:=0;

  if Affaire0 <> '' then
    StatutReduit := Affaire0
  else
  BEGIN
    if bProposition then
      StatutReduit := 'P'
    else
      StatutReduit := 'A';
  END;

  CodeAffaire:=trim(Copy (Affaire,1,15));   // supression de l'avenant

  // Attention a voir avec la gestion des avenants / select
  If (CodeAffaire <>'') Then
    Begin
    sReq := 'SELECT AFF_AFFAIRE,AFF_TIERS FROM AFFAIRE WHERE AFF_AFFAIRE LIKE "'+CodeAffaire+'%"';
    If (Tiers <>'') Then sReq:=sReq + ' AND AFF_TIERS="'+Tiers+'"';
    If Not bAvecAffModele Then sReq:=sReq + ' AND AFF_MODELE="-"';
    if whereParties then
      begin
      if (Affaire1 = '') then
         {$IFDEF BTP}
         BTPCodeAffaireDecoupe(Affaire,Affaire0,Affaire1,Affaire2,Affaire3,Avenant, taCreat,false);
         {$ELSE}
         CodeAffaireDecoupe(Affaire,Affaire0,Affaire1,Affaire2,Affaire3,Avenant, taCreat,false);
         {$ENDIF}
      if Affaire1 <> '' then sReq := sReq + ' AND AFF_AFFAIRE1="'+ Affaire1 + '"';
      if Affaire2 <> '' then sReq := sReq + ' AND AFF_AFFAIRE2="'+ Affaire2 + '"';
      if Affaire3 <> '' then sReq := sReq + ' AND AFF_AFFAIRE3="'+ Affaire3 + '"';
      end;

    if (IdentifiantAffaireRenseigne (Affaire1,Affaire2,Affaire3)) And (Avenant <> '') then
        sReq := sReq + ' AND AFF_AVENANT="'+ Avenant + '"';

// Modif PL le 08/01/02 pour optimiser la requête de recherche d'existance de l'affaire
    Q:=OpenSQL(sReq,True,2) ;
//    Q:=OpenSQL(sReq,True) ;
// Fin modif PL 08/01/02

    // si plus d'une affaire stop le fetch sur les enreg du query ...
    While Not(Q.EOF) And (Result<2) do BEGIN Inc(Result); Q.Next; END;

    // Si 1 affaire
    If (Result=1) then
        Begin
        Q.First;
        Affaire:=Q.FindField('AFF_AFFAIRE').AsString;
        Tiers := Q.FindField('AFF_TIERS').AsString;
        {$IFDEF BTP}
        BTPCodeAffaireDecoupe(Affaire,Affaire0,Affaire1,Affaire2,Affaire3,Avenant, taCreat,false);
        {$ELSE}
        CodeAffaireDecoupe(Affaire,Affaire0,Affaire1,Affaire2,Affaire3,Avenant, taCreat,false);
        {$ENDIF}
        End;
    // Si aucune affaire correspondante
    If (Result=0) then
        Begin
        if (bAvecMessage=true) then
            PGIBoxAf('Aucune affaire correspondante',TitreHalley);
        AFFAIRE2:=''; AFFAIRE3:=''; AVENANT := '00'; Affaire:= StatutReduit + AFFAIRE1;
        End;
    Ferme(Q);
    End;
End;

Function RechercheAffaire (Aff0,Aff1,Aff2,Aff3,Aff4 : String; EnCreatAff : Boolean):Boolean;
Var stSQL : String;
BEGIN
Result := false;
// En création d'affaire si sous forme de compteur sortie rapide car le code sera forcément unique en création
if (EnCreatAff) and ((VH_GC.CleAffaire.Co1Type ='CPT') or
										 (VH_GC.CleAffaire.Co2Type ='CPT') or
                     (VH_GC.CleAffaire.Co3Type ='CPT')) then Exit;

if Aff0 = '' then Aff0:='A';
if Aff4 = '' then Aff4:='00';
stSQL :='SELECT AFF_AFFAIRE FROM AFFAIRE WHERE AFF_AFFAIRE0="'+ Aff0 + '" AND AFF_AVENANT="'+Aff4+'"';
if (Aff1<>'') then stSQL := stSQL + ' AND AFF_AFFAIRE1="' + Aff1 + '"';
if (Aff2<>'') And (VH_GC.CleAffaire.NbPartie > 1) then stSQL := stSQL + ' AND AFF_AFFAIRE2="' + Aff2 + '"';
if (Aff3<>'') And (VH_GC.CleAffaire.NbPartie > 2) then stSQL := stSQL + ' AND AFF_AFFAIRE3="' + Aff3 + '"';
Result := ExisteSQL(stSQL);
END;

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 11/02/2000
Modifié le ... : 11/02/2000
Description .. : Passage automatique sur la partie suivante du code affaire
Mots clefs ... : AFFAIRE;CODE AFFAIRE
*****************************************************************}
Procedure GoSuiteCodeAffaire(PartEnc,PartSuite : THCritMaskEDIT; NumPartieEnc: Integer);
Var Lng : integer;
    Visible : Boolean;
Begin
Visible := True; Lng := 0;

If Not((NumPartieEnc = 1) or (NumPartieEnc = 2)) Then exit;
    case NumPartieEnc of
       1: Begin Lng:= VH_GC.CleAffaire.Co1Lng; Visible:=VH_GC.CleAffaire.Co2Visible; End;
       2: Begin Lng:= VH_GC.CleAffaire.Co2Lng; Visible:=VH_GC.CleAffaire.Co3Visible; End;
       End;
If (Length(PartEnc.Text)=Lng) Then
    Begin
    // Test de la zone + focus partie suivante
    If ((VH_GC.CleAffaire.NbPartie > NumPartieEnc) And (Visible)) Then
       if PartSuite.enabled = True then PartSuite.SetFocus;
    End;
End;


{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 02/02/2000
Modifié le ... : 02/02/2000
Description .. : Modifcation dynamique des colonnes d'un grid en fonction du paramétrage du code affaire 
Mots clefs ... : AFFAIRE;CODE AFFAIRE
*****************************************************************}
{$IFDEF EAGLCLIENT}
Procedure TraiteNomAffaireGrid ( Gr : THGrid; Nom : string; Col : integer);
Var  Lg,ind,NumPart : integer;
     NomChamp,Libelle : string;
     Visible : Boolean;
begin
if nom = '' then Exit;
lg := Length(Nom); ind:= Pos ('_',Nom);
if (lg =0) or (ind = 0) then  exit;
Nomchamp := Copy (Nom, ind+1,lg-ind);
Visible:=False; //mcd 14/02/03
If ((Nomchamp = 'AFFAIRE1') Or (Nomchamp = 'AFFAIRE2') Or (Nomchamp = 'AFFAIRE3')) Then
  Begin
  NumPart := StrToInt(Copy(NomChamp, 8,1));
  Case NumPart Of
      1 : Begin Visible:=True; Libelle:=VH_GC.CleAffaire.Co1Lib;  End;
      2 : Begin Visible:=VH_GC.CleAffaire.Co2Visible; Libelle:=VH_GC.CleAffaire.Co2Lib; End;
      3 : Begin Visible:=VH_GC.CleAffaire.Co3Visible; Libelle:=VH_GC.CleAffaire.Co3Lib; End;
      End;
  if not(Visible) then Gr.colwidths[Col]:=0;
  Gr.Cells[Col,0]:=Libelle;
  End;
if ((Nomchamp='AVENANT') and Not(VH_GC.CleAffaire.GestionAvenant)) then Gr.colwidths[Col]:=0;
end;

Procedure ModifColAffaireGrid (Gr : THGrid; TobQ : TOB );
Var  i,PosFr   : integer;
     Nom,stSQL : string;
begin
if (TOBQ = Nil) then Exit;
i := 0;
if TobQ.FieldCount = 0 then
   begin   // Aucun enregistrement trouvé ...
   stSQL := TOBQ.SQL.Text ; // ont changé dans UTob : String en TStringList...
   if stSQL = '' then exit;
   PosFr := Pos ('FROM',stSQL); if PosFr = 0 then Exit;
   stSQL := Copy (stSQL,8,PosFr-8);
   Nom:=(Trim(ReadTokenPipe(stSQL,',')));
   While (Nom <>'') do
      begin
      TraiteNomAffaireGrid (Gr,Nom,i);
      Nom:=(Trim(ReadTokenPipe(stSQL,',')));
      Inc(i); 
      end;
   end
else
   begin
   for i:=0 to TobQ.FieldCount-1 do
      BEGIN
      Nom := TOBQ.Fields[i].FieldName;
      TraiteNomAffaireGrid (Gr,Nom,i);
      END;
   end;
end;

{$ELSE}
{$IFNDEF ERADIO}
Procedure ModifColAffaireGrid (Gr : TfMul );
var i, NumPart,ind,lg : integer;
    visible : Boolean;
    Libelle,Nomchamp,StCHamp,StChamp2 : String;
Begin
// If Not(Gr is THDBGrid) Then Exit;
if (Gr.Fliste.Columns.Count = 1) then Exit ;

Visible :=False;

For i:=0 to Gr.fliste.Columns.Count-1 do
    Begin
    StChamp := Gr.Q.FormuleQ.GetFormule(Gr.FListe.Columns[i].FieldName); //21/10/2005 on récupère le nom total du champ
    StChamp2 := Gr.FListe.Columns[i].FieldName;
     //mcd 21/10/2005 doit pouvoir être améliore.. voir source XP dans public MCD
    lg := Length(Gr.Fliste.Columns[i].FieldName);
    ind := Pos ('_',Gr.Fliste.Columns[i].FieldName);
    // mcd 30/11/01 car sort si libellé au lieu code !!! if ((ind = 0) or (lg =0)) then  exit;
    if ((ind = 0) or (lg =0)) then  continue;
    if (Gr.Fliste.Columns[i].Field)=Nil then Continue; //mcd 25/10/2005 parfois à Nil et provoque accesvio ..
    Nomchamp := Copy (Gr.Fliste.Columns[i].FieldName,ind+1, lg-ind);
    If ((Nomchamp = 'AFFAIRE1') Or (Nomchamp = 'AFFAIRE2') Or (Nomchamp = 'AFFAIRE3')) Then
        Begin
        NumPart := StrToInt(Copy(NomChamp, 8,1));
        Case NumPart Of
             1 : Begin
                 Visible:=True;
                 Libelle:=VH_GC.CleAffaire.Co1Lib;
                 lg := VH_GC.CleAffaire.Co1lng;
                 End;
             2 : Begin
                 Visible:=VH_GC.CleAffaire.Co2Visible;
                 Libelle:=VH_GC.CleAffaire.Co2Lib;
                 lg := VH_GC.CleAffaire.Co2lng;
                 End;
             3 : Begin
                 Visible:=VH_GC.CleAffaire.Co3Visible;
                 Libelle:=VH_GC.CleAffaire.Co3Lib;
                 lg := VH_GC.CleAffaire.Co3lng;
                 End;
             End;
{$IFNDEF AGL581153}
        Gr.SetVisibleColumn(StChamp,visible) ; //mcd 21/10/2005
{$else}
        Gr.SetVisibleColumn(i,visible) ; //mcd 16/09/05
{$endif}

        // mcd 25/07/02 mis en majuscule pour éviter pb de frappe min/maj
        // mcd 21/10/2005 ?? à priori on test le libellé original , et on ne le fait pas systématiquement, car l'utilisatuer a pu change le libellé
        If ((AnsiUppercase(Copy(Gr.Fliste.Columns[i].Field.DisplayLabel, 1, 6)) = 'PARTIE'))or
         ((Copy(Gr.Fliste.Columns[i].Field.DisplayLabel, 1,19) = 'Code affaire partie'))or            // mcd 30/05/02 dans GL, les libelles ne sont pas les mêmes que dan sles autres tables !!!!!
         ((Copy(Gr.Fliste.Columns[i].Field.DisplayLabel, 1,19) = 'Code Affaire partie'))Then            // mcd 30/05/02 dans Gp , les libelles ne sont pas les mêmes que dan sles autres tables !!!!!
{$IFNDEF AGL581153}
             GR.SetDisplayLabel (stChamp,Libelle);
{$else}
             GR.SetDisplayLabel (i,Libelle);
{$endif}
             //mcd 16/09/2005 agl 148 Gr.columns[i].Field.DisplayLabel:=Libelle ;   // valeur par défaut non modifiée par l'utilisateur
            // Attention dans les listes bien appeler les découpages du code affaire partie 1 , 2, 3
        End
    else if ctxscot in v_pgi.pgicontexte then
          begin
{$IFNDEF AGL581153}
          Gr.SetDisplayLabel (StChamp,TraduitGA(Gr.Fliste.columns[i].Field.DisplayLabel));
{$else}
          Gr.SetDisplayLabel (i,TraduitGA(Gr.Fliste.columns[i].Field.DisplayLabel));
{$endif}
{$IFNDEF BTP}
    // Gestion de la colonne Avenant
    end else if (Nomchamp = 'AVENANT') then
        BEGIN
{$ifndef AGL581153}
        if Not (VH_GC.CleAffaire.GestionAvenant) then  Gr.SetVisibleColumn(stChamp, False);
{$else}
        if Not (VH_GC.CleAffaire.GestionAvenant) then  Gr.SetVisibleColumn(i, False);
{$endif}

        END;
{$ENDIF}
		end else if (copy(NomChamp,1,8) = 'LIBREAFF') or
    						(copy(NomChamp,1,8) = 'VALLIBRE') or
    						(copy(NomChamp,1,9) = 'BOOLLIBRE') or
                (copy(NomChamp,1,9) = 'DATELIBRE') then
      begin
      		GCTitreZoneLibre(stChamp,Libelle);
{$IFNDEF AGL581153}
          Gr.SetDisplayLabel (StChamp,Libelle);
{$else}
          Gr.SetDisplayLabel (i,Libelle);
{$endif}
      end;
    End;
FormatExerciceGrid(Gr.Fliste);
End;
{$ENDIF !ERADIO}
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : Alain Buchet
Créé le ...... : 10/07/2002
Modifié le ... : 10/07/2002
Description .. : Initialise les colonnes d'un Thgrid en fonction du paramétrage du code affaire
Mots clefs ... : AFFAIRE;CODEAFFAIRE;
*****************************************************************}

Procedure ColCodeAffaireGrid (GS : THGRID;LesColonnes:string);
var StCol,Colonne,Champ :string ;
    X,NCol,ColAff,ColAff0,ColAff1,ColAff2,ColAff3,ColAven,ColTiers :integer;
begin
  StCol:= LesColonnes ;
  NCol:= 0; ColAff0 :=-1; ColAff :=-1; ColAff1 :=-1; ColAff2 :=-1; ColAff3 :=-1;
  ColAven :=-1;ColTiers :=-1;
  Colonne:=(Trim(ReadTokenSt(StCol)));
  While (Colonne <>'') do
  begin
    X:=pos('_',Colonne);
    inc(NCol);
    if x<>0 then
    begin
      Champ:=Copy (Colonne,X+1,length(Colonne)-X);
      if Champ = 'AFFAIRE'  then ColAff  := NCol;
      if Champ = 'AFFAIRE0' then ColAff0 := NCol;
      if Champ = 'AFFAIRE1' then ColAff1 := NCol;
      if Champ = 'AFFAIRE2' then ColAff2 := NCol;
      if Champ = 'AFFAIRE3' then ColAff3 := NCol;
      if Champ = 'AVENANT'  then ColAven := NCol;
      if Champ = 'TIERS'    then ColTiers := NCol;
    end;
    Colonne:=(Trim(ReadTokenSt(StCol)));
  end;

  if (ColAff > 0) then GS.ColWidths[ColAff] := 0;
  if (ColAff0 > 0) then GS.ColWidths[ColAff0] := 0;
  if (ColAff1 > 0) then
  begin
    GS.Cells[ColAff1,0]:= VH_GC.CleAffaire.Co1Lib;
    GS.ColWidths[ColAff1]:= 60;
  end;
  if (ColAff2 > 0) then
  begin
    if VH_GC.CleAffaire.Co2Visible then  GS.ColWidths[ColAff2]:= 80
    else GS.ColWidths[ColAff2]:= 0;
    GS.Cells[ColAff2,0]:= VH_GC.CleAffaire.Co2Lib;
  end;
  if (ColAff3 > 0) then
  begin
    if VH_GC.CleAffaire.Co3Visible then  GS.ColWidths[ColAff3]:= 80
    else GS.ColWidths[ColAff3]:= 0;
    GS.Cells[ColAff3,0]:= VH_GC.CleAffaire.Co3Lib;
  end;
  if (ColAven > 0) then
  begin
    if VH_GC.CleAffaire.GestionAvenant then  GS.ColWidths[ColAven]:= 60
    else GS.ColWidths[ColAven]:= 0;
    GS.Cells[ColAven,0]:= 'Avenant';
  end;
  if (ColTiers > 0) then
  begin
    GS.Cells[ColTiers,0]:= 'Client';
    GS.ColWidths[ColTiers]:= 80;
  end;
end;


    //mcd 08/08/00 mise de part0 en affaire0 en THEdit !!!
Function ChargeAffaire(AFFAIRE,AFFAIRE1,AFFAIRE2,AFFAIRE3,AVENANT,TIERS,Affaire0:THCritMaskEdit; TestCle:Boolean; Action : TActionFiche; TobAffaire,TobAffaires : Tob; bAvecMessage, bAvecAffModele,BProposition:boolean) : integer;
Var Part1,Part2,Part3,CodeAvenant,Part0 : string;
    NbAff:Integer;
Begin
NbAff:=1;
Result := 0;
if affaire0 <>NIl then Part0:=Affaire0.text;         //mcd 08/08/00
if (AFFAIRE1.Text='') and (AFFAIRE2.Text='') and (AFFAIRE3.Text='') then
    BEGIN
    TestCle := false;
    NbAff:=0;
    END;
If TestCle Then NbAff:=TestCleAffaire(AFFAIRE,AFFAIRE1,AFFAIRE2,AFFAIRE3,AVENANT,TIERS,Part0,bAvecMessage,bAvecAffModele,bProposition);

If NbAff=1 Then
    Begin
    if Not AFRemplirTOBAffaire(AFFAIRE.Text,TobAffaire,TobAffaires) then
        AFFAIRE.Text:=TOBAffaire.GetValue('AFF_AFFAIRE')
    Else
        Begin
        result := 1;
        End;
    End
else result := NbAff;
{$IFDEF BTP}
//if copy(AFFAIRE.Text,1,1) = 'W' Then
//	 CodeAppelDecoupe(AFFAIRE.Text, Part0,Part1, Part2, Part3,CodeAvenant)
//else
	 BTPCodeAffaireDecoupe(AFFAIRE.Text, Part0,Part1, Part2, Part3,CodeAvenant, Action, False);
{$ELSE}
	 CodeAffaireDecoupe(AFFAIRE.Text, Part0,Part1, Part2, Part3,CodeAvenant, Action, False);
{$ENDIF}
AFFAIRE1.text := Part1;
AFFAIRE2.text := Part2;
AFFAIRE3.text := Part3;
Affaire0.text:=Part0;   //mcd 08/08/00
if Avenant <> Nil then Avenant.Text := CodeAvenant;
End;

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 10/07/2000
Modifié le ... : 10/07/2000
Description .. : La dernière partie visible est renseignée test sur le code 
Suite ........ : avenant possible
Mots clefs ... : AFFAIRE;CODEAFFAIRE;
*****************************************************************}
Function IdentifiantAffaireRenseigne (Part1,Part2,Part3 : string): Boolean;
BEGIN
Result := False;
if (VH_GC.CleAffaire.Co3Visible) then
    BEGIN
    if Part3 <> '' then result := true;
    END else
if (VH_GC.CleAffaire.Co2Visible) then
    BEGIN
    if Part2 <> '' then result := true;
    END
else   // Test sur la partie 1
    BEGIN
    if Part1 <> '' then result := true;
    END;
END;

Procedure RAZIdentifiantAffaire (Var Part1,Part2,Part3 : string);
BEGIN
if Not(VH_GC.CleAffaire.Co3Visible) and (Part3 <> '') then Part3 := '';
if Not(VH_GC.CleAffaire.Co2Visible) and (Part2 <> '') then Part2 := '';
END;

//**********************************************************************************
//**************** Gestion des paramètres du code affaire    ************
//**********************************************************************************

Procedure InitParamCodeAffaire;
BEGIN
SetParamsoc('SO_AFFCODENBPARTIE',3);
SetParamsoc('SO_AFFCO1LNG',3);
SetParamsoc('SO_AFFCO1TYPE','LIS');
SetParamsoc('SO_AFFCO2TYPE','SAI');
SetParamsoc('SO_AFFCO3TYPE','CPT');
SetParamsoc('SO_AFFCO1VALEUR','');
SetParamsoc('SO_AFFCO2VALEUR','');

SetParamsoc('SO_AFFCO1VALEURPRO','');
SetParamsoc('SO_AFFCO2VALEURPRO','');
SetParamsoc('SO_AFFCO2ACT' ,'X');
SetParamsoc('SO_AFFCO3ACT','X');
SetParamsoc('SO_AFFCO2VISIBLE','X');
SetParamsoc('SO_AFFCO3VISIBLE','-');
SetParamsoc('SO_AFFPRODIFFERENT','X');
SetParamsoc('SO_AFFCO1LIB','Mission');
SetParamsoc('SO_AFFCO2LIB','Exercice');
SetParamsoc('SO_AFFCO3LIB','Compteur');
ModifParamCodeAffaire ('A'); // Par défaut exercice au format millésime année seule
SetParamsoc ('SO_AFFORMATEXER','A');
END;

Procedure ModifParamCodeAffaire ( TypeExercice : String);
Var LngPart2, LngPart3 : integer ;
BEGIN
  // mcd 18/09/01 if Not(ctxscot in v_PGI.PGIContexte) then Exit;
if (TypeExercice='AUC') then Exit;
if TypeExercice = 'A'  then LngPart2 := 4 else
if TypeExercice = 'AA' then LngPart2 := 8 else
if TypeExercice = 'AM' then LngPart2 := 6 else  LngPart2 := 4;
LngPart3 := 14 - 3 - LngPart2;
SetParamsoc('SO_AFFCO2LNG',LngPart2);
SetParamsoc('SO_AFFCO3LNG',LngPart3);
SetParamsoc('SO_AFFCO3VALEUR',Copy ('0000000',1,Lngpart3));
SetParamsoc('SO_AFFCO3VALEURPRO',Copy ('0000000',1,Lngpart3));
END;

Procedure InitParpieceAffaire;
BEGIN
if ExisteSQL('SELECT GPP_NATUREPIECEG FROM PARPIECE WHERE GPP_NATUREPIECEG="FRE"') then Exit;
InitUneParpieceAffaire ('FRE');
{$IFDEF EAGLCLIENT}    //mcd 28/10/2005
    AvertirCacheServer('PARPIECE') ;
{$ENDIF}
  // mcd 12/03/01 pour recharger tob peice
if VH_GC.TOBParPiece<>Nil then VH_GC.TOBParPiece.Free ;
VH_GC.TOBParPiece:=TOB.Create('',Nil,-1) ;
VH_GC.TOBParPiece.LoadDetailDB('PARPIECE','','',Nil,True) ;

InitSOucheAffaire();
END;

Procedure ModifLesParpieceAffaire;
BEGIN
ModifUneParPiece('AFF');
ModifUneParPiece('DAP');
ModifUneParPiece('PAF');
ModifUneParPiece('FAC');
ModifUneParPiece('AVC');
ModifUneParPiece('FPR');
ModifUneParPiece('APR');
//mcd 02/07/03 ??? ModifUneParPiece('SIM');
ModifUneParPiece('FSI');
ModifUneParPiece('FF');
ModifUneParPiece('AF'); //mcd 02/07/2003
ModifUneParPiece('FRE'); //mcd 29/07/2003
END;

Procedure InitUneParpieceAffaire(TypePP : string);
Var TOBPP : TOB;
    Lib,Souche,NatureSuivante : string;
BEGIN
TOBPP := TOB.Create ('PARPIECE',Nil,-1);
try
TOBPP.PutValue ('GPP_SENSPIECE','SOR');
if TypePP = '' then exit;

TOBPP.PutValue ('GPP_ACOMPTE','X');
TOBPP.PutValue ('GPP_HISTORIQUE', 'X');
TOBPP.PutValue ('GPP_ENCOURS','-');
if TypePP = 'APR' then
    BEGIN
    TOBPP.PutValue ('GPP_ACOMPTE','-'); //mcd 04/04/02
    TOBPP.PutValue ('GPP_HISTORIQUE', '-');
    TOBPP.PutValue ('GPP_ENCOURS','X');
    END else
if TypePP = 'FPR' then
    BEGIN
    Lib := 'Facture client provisoire';
    NatureSuivante := 'FPR;FAC;';
    Souche := '';
    TOBPP.PutValue ('GPP_ACOMPTE','-'); //mcd 04/04/02
    TOBPP.PutValue ('GPP_HISTORIQUE', '-');
    TOBPP.PutValue ('GPP_ENCOURS','X');
    END else
if TypePP = 'FRE' then
    BEGIN
    TOBPP.PutValue ('GPP_ENCOURS','X');
    Lib := 'Reprise facturation externe';
    NatureSuivante := 'FRE;';
    TOBPP.PutValue ('GPP_SOUCHE','GRE');
    TOBPP.PutValue ('GPP_HISTORIQUE','X');
    TOBPP.PutValue ('GPP_NBEXEMPLAIRE',1);
    TOBPP.PutValue ('GPP_contextes','AFF;GC;'); //mcd 13/10/2005
    TOBPP.PutValue ('GPP_LISTESAISIE','GCSAISIEFAC');
    TOBPP.PutValue ('GPP_LISTEAFFAIRE','AFSAISIEFAC');
    TOBPP.PutValue ('GPP_VENTEACHAT','VEN');
    // mcd 17/11/04 TOBPP.PutValue ('GPP_IMPMODELE','AFC');
    TOBPP.PutValue ('GPP_ACTIONFINI','TRA');
    TOBPP.PutValue ('GPP_SENSPIECE','MIX');
    TOBPP.PutValue ('GPP_AFAFFECTTB','FAC');
    TOBPP.PutValue ('GPP_NIVEAUPARAM','EDI');
    TOBPP.PutValue ('GPP_MODEGROUPEPORT','CUM');
    END else
if TypePP = 'FSI' then
    BEGIN
    Lib := 'Facture de simulation';
    NatureSuivante := 'FSI';
    Souche := '';
    END else Exit;

TOBPP.PutValue ('GPP_MAJINFOTIERS','X');  //MCD 27/10/03 il faut que ce soit à X pour OK
TOBPP.PutValue ('GPP_NATUREPIECEG',TypePP);
TOBPP.PutValue ('GPP_LIBELLE',Lib);
TOBPP.PutValue ('GPP_NATURETIERS','CLI;');
TOBPP.PutValue ('GPP_RECALCULPRIX','X');
TOBPP.PutValue ('GPP_CONDITIONTARIF','X');
TOBPP.PutValue ('GPP_ENCOURS','-');
TOBPP.PutValue ('GPP_LIENAFFAIRE','X');
TOBPP.PutValue ('GPP_RELIQUAT','-');
TOBPP.PutValue ('GPP_GEREECHEANCE','DEM');
TOBPP.PutValue ('GPP_TYPEECRCPTA','RIE');
TOBPP.PutValue ('GPP_NATURESUIVANTE',NatureSuivante);
TOBPP.PutValue ('GPP_APPELPRIX','PUH');
TOBPP.PutValue ('GPP_COMPANALLIGNE','SAN');
TOBPP.PutValue ('GPP_COMPANALPIED','SAN');
TOBPP.PutValue ('GPP_CALCRUPTURE','AUC');
TOBPP.PutValue ('GPP_VALLIBART1','AUC');
TOBPP.PutValue ('GPP_VALLIBART2','AUC');
TOBPP.PutValue ('GPP_VALLIBART3','AUC');
TOBPP.PutValue ('GPP_DATELIBART1','AUC');
TOBPP.PutValue ('GPP_DATELIBART2','AUC');
TOBPP.PutValue ('GPP_DATELIBART3','AUC');
TOBPP.PutValue ('GPP_DATELIBTIERS1','AUC');
TOBPP.PutValue ('GPP_DATELIBTIERS2','AUC');
TOBPP.PutValue ('GPP_DATELIBTIERS3','AUC');
TOBPP.PutValue ('GPP_VALLIBCOM1','AUC');
TOBPP.PutValue ('GPP_VALLIBCOM2','AUC');
TOBPP.PutValue ('GPP_VALLIBCOM3','AUC');
TOBPP.PutValue ('GPP_DATELIBCOM1','AUC');
TOBPP.PutValue ('GPP_DATELIBCOM2','AUC');
TOBPP.PutValue ('GPP_DATELIBCOM3','AUC');
TOBPP.PutValue ('GPP_VALLIBTIERS1','AUC');
TOBPP.PutValue ('GPP_VALLIBTIERS2','AUC');
TOBPP.PutValue ('GPP_VALLIBTIERS3','AUC');
TOBPP.PutValue ('GPP_TYPEARTICLE',PlusTypeArticleText);
// mcd 07/03/02 TOBPP.PutValue ('GPP_TYPEARTICLE','FRA;MAR;PRE');
TOBPP.PutValue ('GPP_EDITIONNOMEN','AUC');
TOBPP.PutValue ('GPP_PRIORECHART1','ART');
TOBPP.PutValue ('GPP_PRIORECHART2','REF');
TOBPP.PutValue ('GPP_PRIORECHART3','BAR');
TOBPP.PutValue ('GPP_IFL1','001');
TOBPP.PutValue ('GPP_IFL2','016');
TOBPP.PutValue ('GPP_IFL3','017');
TOBPP.PutValue ('GPP_IFL4','019');
TOBPP.PutValue ('GPP_IFL5','003');
TOBPP.PutValue ('GPP_TAXE','X');
  //mcd 17/11/04 ajoput nouveaux champs
TOBPP.PutValue ('GPP_INSERTLIG','X');
TOBPP.PutValue ('GPP_ACHATACTIVITE','-');
TOBPP.PutValue ('GPP_AFFPIECETABLE','-');
TOBPP.PutValue ('GPP_APERCUAVETIQ','-');
TOBPP.PutValue ('GPP_APERCUAVIMP','-');
TOBPP.PutValue ('GPP_APPLICRG','-');
TOBPP.PutValue ('GPP_ARTFOURPRIN','-');
TOBPP.PutValue ('GPP_ARTSTOCK','-');
TOBPP.PutValue ('GPP_BLOBART','-');
TOBPP.PutValue ('GPP_BLOBTIERS','-');
TOBPP.PutValue ('GPP_CONTRECHART1','-');
TOBPP.PutValue ('GPP_CONTRECHART2','-');
TOBPP.PutValue ('GPP_CONTRECHART3','-');
TOBPP.PutValue ('GPP_CODEPIECEOBL1','-');
TOBPP.PutValue ('GPP_CODEPIECEOBL2','-');
TOBPP.PutValue ('GPP_CODEPIECEOBL3','-');
TOBPP.PutValue ('GPP_CONTREMARQUE','-');
TOBPP.PutValue ('GPP_CONTREMREF','-');
TOBPP.PutValue ('GPP_CPTCENTRAL','-');
TOBPP.PutValue ('GPP_CUMULART1','-');
TOBPP.PutValue ('GPP_CUMULART2','-');
TOBPP.PutValue ('GPP_CUMULART3','-');
TOBPP.PutValue ('GPP_CUMULCOM1','-');
TOBPP.PutValue ('GPP_CUMULCOM2','-');
TOBPP.PutValue ('GPP_CUMULCOM3','-');
TOBPP.PutValue ('GPP_CUMULTIERS1','-');
TOBPP.PutValue ('GPP_CUMULTIERS2','-');
TOBPP.PutValue ('GPP_CUMULTIERS3','-');
TOBPP.PutValue ('GPP_ECLATEAFFAIRE','-');
TOBPP.PutValue ('GPP_ECLATEDOMAINE','-');
TOBPP.PutValue ('GPP_ESTAVOIR','-');
TOBPP.PutValue ('GPP_FORCERUPTURE','-');
TOBPP.PutValue ('GPP_IMPETIQ','-');
TOBPP.PutValue ('GPP_LIENTACHE','-');
TOBPP.PutValue ('GPP_LOT','-');
TOBPP.PutValue ('GPP_MASQUERNATURE','-');
TOBPP.PutValue ('GPP_MODIFCOUT','-');
TOBPP.PutValue ('GPP_MULTIGRILLE','-');
TOBPP.PutValue ('GPP_NUMEROSERIE','-');
TOBPP.PutValue ('GPP_OBJETDIM','-');
TOBPP.PutValue ('GPP_OBLIGEREGLE','-');
TOBPP.PutValue ('GPP_OUVREAUTOPORT','-');
TOBPP.PutValue ('GPP_PARAMDIM','-');
TOBPP.PutValue ('GPP_PARAMGRILLEDIM','-');
TOBPP.PutValue ('GPP_PIECEPILOTE','-');
TOBPP.PutValue ('GPP_PREVUAFFAIRE','-');
TOBPP.PutValue ('GPP_PROCLI','-');
TOBPP.PutValue ('GPP_TRSFACHAT','-');
TOBPP.PutValue ('GPP_TRSFVENTE','-');
TOBPP.PutValue ('GPP_TYPEDIMOBLI1','-');
TOBPP.PutValue ('GPP_TYPEDIMOBLI2','-');
TOBPP.PutValue ('GPP_TYPEDIMOBLI3','-');
TOBPP.PutValue ('GPP_TYPEDIMOBLI4','-');
TOBPP.PutValue ('GPP_TYPEDIMOBLI5','-');
TOBPP.PutValue ('GPP_VALMODELE','-');
TOBPP.PutValue ('GPP_VISA','-');
TOBPP.PutValue ('GPP_IMPAUTOBESCBN','-');
TOBPP.PutValue ('GPP_IMPAUTOETATCBN','-');
TOBPP.PutValue ('GPP_IMPBESOIN','-');
TOBPP.PutValue ('GPP_PIECEEDI','-');
TOBPP.PutValue ('GPP_PRIXNULOK','-');
TOBPP.PutValue ('GPP_PILOTEORDRE','-');
TOBPP.PutValue ('GPP_REFEXTCTRL','000');
TOBPP.PutValue ('GPP_REFINTCTRL','000');
TOBPP.PutValue ('GPP_REFINTEXT','INT');
TOBPP.PutValue('GPP_IMPIMMEDIATE', '-');
TOBPP.PutValue('GPP_IMPMODELE',VH_GC.GCImpModeleDefaut );
TOBPP.PutValue ('GPP_COMPSTOCKLIGNE','SAN');
TOBPP.PutValue ('GPP_COMPSTOCKPIED','SAN');
TOBPP.PutValue ('GPP_REGROUPCPTA','AUC');
TOBPP.PutValue ('GPP_CONTEXTES','AFF;');
TOBPP.PutValue ('GPP_RECUPPRE','PRE');
TOBPP.PutValue ('GPP_TYPEPRESDOC','DEF');
TOBPP.PutValue ('GPP_TARIFMODULE','201');
TOBPP.PutValue ('GPP_STKQUALIFMVT','SVE');
TOBPP.PutValue ('GPP_VALLIBCOM1','AUC');
TOBPP.PutValue ('GPP_VALLIBCOM2','AUC');
TOBPP.PutValue ('GPP_VALLIBCOM3','AUC');
TOBPP.PutValue ('GPP_DATELIBCOM1','AUC');
TOBPP.PutValue ('GPP_DATELIBCOM2','AUC');
TOBPP.PutValue ('GPP_DATELIBCOM3','AUC');
TOBPP.PutValue ('GPP_VALLIBART1','AUC');
TOBPP.PutValue ('GPP_VALLIBART2','AUC');
TOBPP.PutValue ('GPP_VALLIBART3','AUC');
TOBPP.PutValue ('GPP_DATELIBART1','AUC');
TOBPP.PutValue ('GPP_DATELIBART2','AUC');
TOBPP.PutValue ('GPP_DATELIBART3','AUC');
TOBPP.PutValue ('GPP_DIMSAISIE','TOU');

  // fin maj mcd 17/11/2004
TOBPP.InsertOrUpdateDB  ;
finally
TOBPP.free;
end;
END;

Procedure ModifUneParPiece(TypePP : string);
Var TOBPP : TOB;
    Req : String;
    Q : TQuery;
BEGIN
TOBPP := TOB.Create ('PARPIECE',Nil,-1);
try
   // il faut bien prendre toutes les info de ParPiece, tout peut être utile
Req := 'SELECT * from PARPIECE where GPP_NATUREPIECEG="'+ TypePP + '"';
Q:=OpenSQL(Req,True,-1,'',true);
If (Not Q.EOF) then TOBPP.SelectDB('',Q) Else exit;
Ferme(Q);
// par defaut sur toutes les natures
TOBPP.PutValue ('GPP_IFL1','001');
TOBPP.PutValue ('GPP_IFL2','016');
TOBPP.PutValue ('GPP_IFL3','017');
TOBPP.PutValue ('GPP_IFL4','019');
TOBPP.PutValue ('GPP_IFL5','003');
TOBPP.PutValue ('GPP_IFL6','');
TOBPP.PutValue ('GPP_IFL7','');
TOBPP.PutValue ('GPP_IFL8','');
TOBPP.PutValue ('GPP_MAJINFOTIERS','X');  //MCD 20/10/03 il faut que ce soit à X pour OK
TOBPP.PutValue ('GPP_TYPEARTICLE',PlusTypeArticleText); // mcd 07/03/02 pour mettre valeur gérée par Produit
// reparamétrages des modèles d'éditions
if (TypePP = 'FAC') then // Sur factures
    BEGIN
    TOBPP.PutValue ('GPP_IMPMODELE','AH1');
    TOBPP.PutValue ('GPP_JOURNALCPTA','VEN');   // cohérence avec journaux comptables créés
    TOBPP.PutValue ('GPP_IMPETAT','AH1'); 
    END
else if (TypePP = 'AVC') then // Sur avoirs
    BEGIN
    TOBPP.PutValue ('GPP_IMPMODELE','AV2');
    TOBPP.PutValue ('GPP_JOURNALCPTA','VEN');   // cohérence avec journaux comptables créés
    TOBPP.PutValue ('GPP_IMPETAT','AH1');
    END
else if (TypePP = 'APR') then // Sur avoirs provisoires
    BEGIN
    TOBPP.PutValue ('GPP_IMPMODELE','AV2');
    TOBPP.PutValue ('GPP_IMPETAT','AH1'); 
    END
else if (TypePP = 'FPR') or (TypePP = 'FSI') then  // Factures porivisoires
    BEGIN
    TOBPP.PutValue ('GPP_IMPMODELE','AHP'); //mcd TOBPP.PutValue ('GPP_IMPMODELECON','AP2');
    TOBPP.PutValue ('GPP_IMPETAT','AH1'); //TOBPP.PutValue ('GPP_IMPETATCON','AH1');
    END
else if (TypePP = 'AFF') or (TypePP = 'PAF') then  // MCD 06/10/03 pour création affaire/provisoire
    BEGIN
    TOBPP.PutValue ('GPP_COMPANALLIGNE','DEM');
    END
else if (TypePP = 'FF')or (TypePP = 'AF') then // fiat à
    BEGIN
    TOBPP.PutValue ('GPP_LISTEAFFAIRE','AFSAISIEFFAC');
    TOBPP.PutValue ('GPP_JOURNALCPTA','ACH');
    TOBPP.PutValue ('GPP_ACHATACTIVITE','X');
    TOBPP.PutValue ('GPP_TYPEARTICLE','MAR'); // mcd 02/07/03
    TOBPP.PutValue ('GPP_LIENAFFAIRE','X');  //mcd 02/07/03
    END;
TOBPP.UpdateDB;
finally
TOBPP.Free;
end;
END;


Procedure InitSoucheAffaire;
Var TOBS : TOB;
BEGIN
     // crée la souche GRE associé à FRE
TOBS := TOB.Create ('SOUCHE',Nil,-1);
try
TOBS.PutValue ('SH_TYPE','GES');
TOBS.PutValue ('SH_SOUCHE','GRE');
TOBS.PutValue ('SH_LIBELLE','Souche pour reprise facture GI/GA');
TOBS.PutValue ('SH_ABREGE','Reprise facture');
TOBS.PutValue ('SH_NUMDEPART',1);
TOBS.PutValue ('SH_NUMDEPARTS',1);
TOBS.PutValue ('SH_NUMDEPARTP',1);
TOBS.PutValue ('SH_SOCIETE',V_PGI.CodeSociete);
TOBS.InsertOrUpdateDB  ;
finally
TOBS.Free;
end;
END;

function PropositionToAffaire(CodeProposition:string; var Aff0, Aff1, Aff2, Aff3, Aff4:string; bGarderCpt,bUpdateCpt : Boolean):string;
var
Pp0,Pp1,Pp2,Pp3,Pp4:string;
begin
Result:='';
if (CodeProposition <> '')  then
{$IFDEF BTP}
    BTPCodeAffaireDecoupe(CodeProposition, Pp0, Pp1, Pp2, Pp3, Pp4, taCreat, false);
{$ELSE}
    CodeAffaireDecoupe(CodeProposition, Pp0, Pp1, Pp2, Pp3, Pp4, taCreat, false);
{$ENDIF}


Aff0:= 'A';  // Partie 0 = affaire
Aff4:= Pp4;  // Partie 4 = 00 ( pas d'avenant)
// Partie 1, Partie 2, Partie 3
Result:=CodeAffaireRegroupe(Aff0, Pp1 ,Pp2 ,Pp3, Aff4, taCreat, True, bGarderCpt,bUpdateCpt);
Aff1:=Pp1;
Aff2:=Pp2;
Aff3:=Pp3;

end;

//**********************************************************************************
//**************** Fonctions de calcul du nombre d'échéances     ************
//**********************************************************************************
{***********A.G.L.***********************************************
Auteur  ...... : G. Merieux

Créé le ...... : 03/02/2000

Modifié le ... :   /  /

Description .. : Recherche par affaire du nbre d'échéance,montant total

Mots clefs ... : ECHEANCE FACTURATION;ECHEANCE AFFAIRE

*****************************************************************}

Procedure  CumulEcheances(CodeAff:string;var zcum:double;var zpou:double;var znbr:integer);
Var QQ : TQuery;
Begin
// PL le 06/03/02 : INDEX 3
//mcd 30/09/03 ajout pour ne pas prendre en compte ech liquidative
QQ:=OpenSQL('SELECT COUNT(AFA_NUMECHE),SUM(AFA_MONTANTECHEDEV),SUM(AFA_POURCENTAGE) FROM FACTAFF WHERE AFA_TYPECHE="NOR" AND AFA_AFFAIRE="'+CodeAff+'" AND AFA_LIQUIDATIVE="-"',TRUE,-1,'',true) ;
if Not QQ.EOF then
   begin
   znbr:=QQ.Fields[0].AsInteger;
   zcum:=QQ.Fields[1].AsFloat;
   zpou:=QQ.Fields[2].AsFloat;
   end;
Ferme(QQ);
End;


function EvaluationNbEcheances(CodeAff, TypeCalcul : string; zinterval: integer; DateDebut , DateFin : TdateTime ):integer;
var
  Rang, RangInitial:integer;
  QQ : TQuery;
  DateEcheance, dDateMax:TDateTime;
  sReq : string;
begin

  Result:=0;
  Rang:=0;
  // Recherche de la derniere date facturee et du rang associe

  // PL le 06/03/02 : INDEX 3
  sReq := 'SELECT AFA_DATEECHE,AFA_NUMECHE FROM FACTAFF Where  AFA_TYPECHE="NOR" AND AFA_AFFAIRE="'+CodeAff+'" AND AFA_ECHEFACT="X" AND AFA_LIQUIDATIVE="-" ORDER BY AFA_DATEECHE DESC';
  // PL le 09/01/02 pour optimiser le nombre d'enregistrement lus en eagl
  QQ := nil;
  Try
    QQ:=OpenSQL(sReq,True,1);
    //QQ:=OpenSQL(sReq,True);
    if Not QQ.EOF then
       begin
       QQ.First;  // classées par ordre décroissant
       dDateMax := QQ.Fields[0].AsDateTime;
       if IsValidDate(datetostr(dDateMax)) and (dDateMax<>0) and (dDateMax<>iDate1900) and (dDateMax<>iDate2099) then
          begin
          if (dDateMax>=DateDebut) then
             DateDebut := PlusDate(dDateMax, zinterval, TypeCalcul);
          Rang:=QQ.Fields[1].AsInteger;
          end;
       end;
  Finally
    Ferme(QQ);
  end;
  result := rang;

  if TypeCalcul ='P' then begin
                result:=1;
                exit; // mcd 04/02/02 si ponctuelle on ne va pas voir les echeances
                end;
  DateEcheance := DateDebut;
  RangInitial := Rang;
  // calcul du nombre d'échéances
  while (DateEcheance <= DateFin)  do
    Begin
      Inc(Rang);
      Inc(Result);
      DateEcheance := PlusDate(DateDebut,(Rang-RangInitial) * zinterval,TypeCalcul);
    End;
end;

Procedure ConvertDevToPivot(DEV : RDEVISE; MtDev : Double;  var MtPivot : double);
BEGIN
  // C.B 18/06/2003 Suppression contrevaleur +mcd supression SisieCOntre
  if (DEV.Code <> V_PGI.DevisePivot) then
    MtPivot := DeviseToPivot(MtDev,DEV.Taux,DEV.Quotite)
  else    MtPivot := MtDev;
END;

Procedure ConvertPivotToDev(DEV : RDEVISE; MtPivot : Double;  var  MtDev : double);
BEGIN
  // C.B 18/06/2003 Suppression contrevaleur + mcd suppression SaisieSontre
  if (DEV.Code <> V_PGI.DevisePivot) then
    MtDev := PivotToDevise(MtPivot,DEV.Taux,DEV.Quotite,DEV.Decimale)
  else MtDev := MtPivot;
END;

Procedure InverseEcheanceContre (CodeAff : String ; SaisieContre : boolean);
Var TobEch, TobDet : TOB;
    ii : integer;
BEGIN
TobEch := Tob.Create ('Liste echeances', Nil,-1);
ChargeEcheances(CodeAff,'','', iDate1900, iDate2099, TobEch);  //gm 08/07/02
if (TobEch.Detail.Count > 0) Then
    Begin
    for ii:=0 To TobEch.Detail.Count-1 do
        Begin
        TobDet := TobEch.Detail[ii];
        // if SaisieContre then Suf:='CON' else Suf:='' ;
        if not SaisieContre then
         begin
         if (TobDet.GetValue('AFA_TYPECHE')='NOR') then
            TobDet.PutValue('AFA_MONTANTECHEDEV', TobDet.GetValue('AFA_MONTANTECHE'))
         else
            begin
            TobDet.PutValue('AFA_BONIMALIDEV',  TobDet.GetValue('AFA_BONIMALI'));
            TobDet.PutValue('AFA_BM1FODEV',     TobDet.GetValue('AFA_BM1FO'));
            TobDet.PutValue('AFA_BM1FRDEV',     TobDet.GetValue('AFA_BM1FR'));
            TobDet.PutValue('AFA_BM2PRDEV',     TobDet.GetValue('AFA_BM2PR'));
            TobDet.PutValue('AFA_BM2FODEV',     TobDet.GetValue('AFA_BM2FO'));
            TobDet.PutValue('AFA_BM2FRDEV',     TobDet.GetValue('AFA_BM2FR'));
            TobDet.PutValue('AFA_AFACTOTDEV',   TobDet.GetValue('AFA_AFACTOT'));
            TobDet.PutValue('AFA_AFACTURERDEV', TobDet.GetValue('AFA_AFACTURER'));
            TobDet.PutValue('AFA_AFACTFRDEV',   TobDet.GetValue('AFA_AFACTFR'));
            TobDet.PutValue('AFA_AFACTFODEV',   TobDet.GetValue('AFA_AFACTFO'));
            TobDet.PutValue('AFA_ACPTEPRDEV',   TobDet.GetValue('AFA_ACPTEPR'));
            TobDet.PutValue('AFA_ACPTEFODEV',   TobDet.GetValue('AFA_ACPTEFO'));
            TobDet.PutValue('AFA_ACPTEFRDEV',   TobDet.GetValue('AFA_ACPTEFR'));
            end;
          end;
        END;
    END;
TobEch.UpdateDB;
TobEch.free;
END;

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 15/12/2000
Modifié le ... : 15/12/2000
Description .. : Création d'une échéance en automatique
Mots clefs ... : FACTAFF;ECHEANCES
*****************************************************************}
function CreerFactAff (TobAffaire : TOB; DateFact : TDateTime;ztyp : string  ): Boolean;
Var iMax, posit : integer;
    QQ : TQuery;
    TobEch : TOB;
    Affaire,st : string;
    Continue : Boolean;
BEGIN
Result := False;
if TobAffaire = Nil then Exit;
if (DateFact = iDate1900) or (DateFact = iDate2099) then
   BEGIN
   PGIBoxAf ('Date d''échéance de facturation incorrecte','Génération d''échéance de facturation');
   Exit;
   END;
Affaire := TOBAffaire.GetValue('AFF_AFFAIRE');
if Affaire = '' then Exit;

if (ztyp = 'NOR') then
  Begin
  {$IFDEF BTP}
  if (PGIAskAF('Confirmer-vous la génération d''une échéance de facturation sur l''affaire '+
     BTPCodeAffaireAffiche(Affaire)+' au '+ DateToStr(DateFact),'')<> mrYes) then Exit;
  End;
  {$ELSE}
  if (PGIAskAF('Confirmer-vous la génération d''une échéance de facturation sur l''affaire '+
     CodeAffaireAffiche(Affaire)+' au '+ DateToStr(DateFact),'')<> mrYes) then Exit;
  End;
  {$ENDIF}

// Test si une echéance existe sur cette date
if ExisteSQL('SELECT AFA_NUMECHE FROM FACTAFF WHERE AFA_TYPECHE="'+ztyp+'" AND AFA_AFFAIRE="'+Affaire+'" AND AFA_DATEECHE = "'+usdatetime(DateFact)+'"') then
   BEGIN
   if (PGIAskAF('Attention une échéance existe à cette date, Souhaitez-vous la générer','')= mrYes) then
      continue:=True else continue:=false;
   END
else continue := true;
// Ferme(QQ); gm : inutile ?
if Not(Continue) then Exit;
// Création réelle de l'échéance
// Recup dernier num rang
// attention le select max()  rend toujours QQ.EOF=false !!!!
//imax := 0;
// PL le 06/03/02 : INDEX 3
QQ:=OpenSQL('SELECT max(AFA_NUMECHE) FROM FACTAFF WHERE AFA_TYPECHE="'+ztyp+'" AND AFA_AFFAIRE="'+Affaire+'"',TRUE,-1,'',true) ;
iMax:=QQ.Fields[0].AsInteger;
if (imax <> 0) then
	iMax:=QQ.Fields[0].AsInteger+1
else
  begin
  	if (ztyp = 'NOR') then iMax:=1 else iMax := 5001;
  end;

Ferme(QQ);
// Alim TOB
TobEch := TOB.Create('FACTAFF',Nil,-1);

TobEch.PutValue('AFA_DEVISE',TOBAffaire.GetValue('AFF_DEVISE'));
TobEch.PutValue('AFA_TIERS',TOBAffaire.GetValue('AFF_TIERS'));

TobEch.PutValue('AFA_AFFAIRE',Affaire);
TobEch.PutValue('AFA_TYPECHE',ztyp);
TobEch.PutValue('AFA_NUMECHE',IMax);
TobEch.PutValue('AFA_NUMECHEBIS',IMax);  //mcd 15/05/03
TobEch.PutValue('AFA_DATEECHE',DateFact);
   // Traitement du libellé
St := VH_GC.AFFLibfactAff;
Posit := Pos ('**',St);            // **  pour reprendre le numéro d'échéance
If (Posit <> 0) then
   Begin delete(St,Posit,2); Insert(TobEch.GetValue('AFA_NUMECHEBIS'),St,Posit); end;   //mcd 23/05/03 mise ech bis
Posit := Pos ('$$',St);            // $$  pour reprendre la date
If (Posit <> 0) then
   Begin delete(St,Posit,2); Insert(TobEch.GetValue('AFA_DATEECHE'),St,Posit); end;
TobEch.PutValue('AFA_LIBELLEECHE',St);
TobEch.PutValue('AFA_ECHEFACT','-');
TobEch.PutValue('AFA_REPRISEACTIV',TOBAffaire.GetValue('AFF_REPRISEACTIV'));
TobEch.PutValue('AFA_LIQUIDATIVE','-');
/// mcd 03/07/2003 profil alimenter dans tous les cas if (ztyp = 'APP') then  //GISE
TobEch.PutValue('AFA_PROFILGENER',TOBAffaire.GetValue('AFF_PROFILGENER'));

Result := TobEch.InsertDB(Nil);
TobEch.Free;
END;

//**********************************************************************************
//           Chargement TOB des echeances  A facturer  d'une affaire
//              pour une tranche de date
//**********************************************************************************
Procedure ChargeEcheances(CodeAff,CodeGenerauto,typech:string;DateDebut,DateFin:TdateTime;TobEch:Tob; AvecFactures:boolean=false);
Var QQ : Tquery;
    req,zch : string;
Begin

  if (typech = '') then
  // SELECT * : on a vraiment besoin de tous les champs
    req := 'SELECT * FROM FACTAFF Where AFA_AFFAIRE="' + CodeAff + '"'
  else
  if (typech = 'APP') then
  // PL le 06/03/02 : INDEX 3
  // SELECT * : on a vraiment besoin de tous les champs
    req := 'SELECT * FROM FACTAFF Where AFA_TYPECHE="'+typech+'" AND AFA_AFFAIRE="' + CodeAff + '"'
  else
  if (typech = 'NOR') then
      begin
    // gm : dans le cas d'echeance normal, je ne prends que les zones concernnées,
    // Je n'ai pas besoin    de tout le SOUC appreciation
    //mcd 15/05/03 ajout NUMECHEBIS
    zch := 'AFA_TIERS,AFA_AFFAIRE,AFA_NUMECHE,AFA_NUMECHEBIS,AFA_TYPECHE,AFA_MONTANTECHEDEV,AFA_MONTANTECHE,AFA_POURCENTAGE';
    zch := zch+',AFA_DATEECHE,AFA_LIBELLEECHE,AFA_GENERAUTO,AFA_ECHEFACT,AFA_REPRISEACTIV';
    zch := zch+',AFA_PROFILGENER,AFA_NUMPIECE ';
  // PL le 06/03/02 : INDEX 3
    req := 'SELECT '+zch+' FROM FACTAFF Where  AFA_TYPECHE="'+typech+'" AND AFA_AFFAIRE="' + CodeAff + '"';
    end;

  if (codegenerauto <>'') then    // gm 08/07/02
  begin
    req := req + ' and AFA_GENERAUTO = "'+CodeGenerauto+'"';

  end;

  if Not AvecFactures then req := Req + ' and AFA_ECHEFACT = "-"';

  if DateDebut <> iDate1900 then req := Req + ' and AFA_DATEECHE >= "'+usdatetime(DateDebut)+'"';
  if DateFin <> iDate2099 then req := Req + ' and AFA_DATEECHE <= "'+usdatetime(DateFin)+'"';

  QQ := nil;
  Try
    QQ := OpenSQL(Req,True,-1,'',true) ;
    If Not QQ.EOF then TobEch.LoadDetailDB('FACTAFF','','',QQ,True);
  Finally
    Ferme(QQ);
  End;
End;

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 02/02/2000
Modifié le ... : 02/02/2000
Description .. : Calcul automatique du total de l'affaire en fonction des échéances
Mots clefs ... : AFFAIRE;ECHEANCE AFFAIRE
*****************************************************************}
Function CalculTotalAffaire(CodeAff, TypeGener, TypePrevu : String; TotalHTLigne : double; TobEcheances: TOB; var NbEche : integer; var baseEche : Boolean; DateDebGener :TDatetime ) : Double;
Var i : integer;
    CumEche, PourcentEche : double;
    bNbEche :  Boolean;
Begin
 //Result := 0;
 NbEche := 0;
CumEche :=0; PourcentEche :=0;
//bNbEche := False;
baseEche := False;

if TotalHTLigne = 0 then
   BEGIN baseEche :=true; bNbEche := true; END
else   // total ligne <> 0 - calcul basé sur les lignes
   BEGIN
   if TypePrevu = 'GEN' then BEGIN baseEche :=false; bNbEche := true; END
                        else BEGIN baseEche :=false; bNbEche := false; NbEche := 1; END;
   END;
// accélérer les traitement - evite une relecture du factaff ... PA le 11/05/2001
if (DateDebGener =idate1900) then        // sup or (TypeGener = 'MAN') on prend des echeances si elles existent
   BEGIN baseEche :=false; bNbEche := false; NbEche := 1; END;

// Gestion du nombre et du cumul des écheances
if (baseEche) or (bNbEche) then
   BEGIN
   if TobEcheances <> Nil then
      BEGIN
      NbEche := TobEcheances.detail.count;
      if baseEche then
        BEGIN
        for i := 0 to TobEcheances.detail.count-1 do
            CumEche := CumEche + TobEcheances.Detail[i].Getvalue('AFA_MONTANTECHEDEV');
        END;
      END
   else CumulEcheances(CodeAff, CumEche, PourcentEche, NbEche);
   END;

// Calcul du Total HT
if (TypeGener = 'FOR') or (TypeGener = 'ACT') or (TypeGener = 'MAN') then
   BEGIN  // reprise des lignes ou des montants forfaitaires
   if TotalHTLigne <> 0 then Result:= TotalHTLigne * NbEche
                        else Result:= CumEche;
   END
else    // dans les autres cas reprise des lignes uniquement
   BEGIN
   Result:= TotalHTLigne * NbEche;
   END;

End;

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 26/04/2000
Modifié le ... : 26/04/2000
Description .. : Retourne le montant mensuel des échéances d'une affaire
Mots clefs ... : AFFAIRE;ECHEANCE AFFAIRE
*****************************************************************}
Function  EcheanceMoisAffaire (CodeAff : String; DateMois : TDatetime) : Double;
Var QQ : Tquery;
    DateDeb, DateFin : TDateTime;
    stSQL : string;
BEGIN
  Result := 0.0;
  DateDeb := DebutdeMois(DateMois);
  DateFin := FinDeMois(DateMois);
  // PL le 06/03/02 : INDEX 3
  stSQL := 'SELECT AFA_MONTANTECHE,AFA_POURCENTAGE, AFF_GENERAUTO, AFF_TOTALHTDEV FROM FACTAFF FA'+
        ' left outer join AFFAIRE on AFF_AFFAIRE=AFA_AFFAIRE'+
        ' WHERE AFA_TYPECHE="NOR" AND AFA_AFFAIRE="'+CodeAff+'"'+
        ' and AFA_DATEECHE >= "'+usdatetime(DateDeb)+'"'+
        ' and AFA_DATEECHE <= "'+usdatetime(DateFin)+'"';

  QQ := nil;
  Try
    QQ := OpenSQL(stSQL,True,-1,'',true) ;
    If Not QQ.EOF then
        BEGIN
        While not QQ.EOF do
            BEGIN
            if (QQ.FindField('AFF_GENERAUTO').asString = 'POU') Or (QQ.FindField('AFF_GENERAUTO').AsString = 'POT') then
                Result := Result + Arrondi(QQ.FindField('AFA_POURCENTAGE').AsFloat * 0.01 * QQ.FindField('AFF_TOTALHTDEV').AsFloat,V_PGI.OkDecV)
            else
                Result := Result + QQ.FindField('AFA_MONTANTECHE').AsFloat;

            QQ.Next;   // seulement quelques échéances par période...
            END;
        END;
  Finally
    Ferme(QQ);
  End;
END;

//******************************************************************************
//*********************** Génération de pièces sur l'affaire *******************
//******************************************************************************

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 02/02/2000
Modifié le ... : 02/02/2000
Description .. : Création automatique de la pièce (vide) associée à l'affaire
Mots clefs ... : AFFAIRE;PIECE AFFAIRE
*****************************************************************}
Function CreerPieceAffaire(CodeTiers, CodeAffaire, StatutAffaire,Etablissement : String;Var CleDocAffaire : R_CLEDOC; EnHT,SaisieContre : Boolean):Boolean;
Begin
result := true;
FillChar(CleDocAffaire,Sizeof(CleDocAffaire),#0) ;
If (StatutAffaire='AFF') Then
   CleDocAffaire.NaturePiece:=VH_GC.AFNatAffaire
{$IFDEF BTP}
else if (StatutAffaire='INT') Then
   CleDocAffaire.NaturePiece:= 'AFF'
else if (StatutAffaire='APP') Then
   CleDocAffaire.NaturePiece:= 'DAP'
{$ENDIF}
Else
  CleDocAffaire.NaturePiece:=VH_GC.AFNatProposition;
  
if CleDocAffaire.NaturePiece = '' then
    BEGIN
    PgiBoxAf ('Nature de pièce associée non paramétrée','Création de pièce sur affaire');
    Exit;
    END;
CleDocAffaire.DatePiece:=V_PGI.DateEntree;
Result := CreerPieceVide(CleDocAffaire, CodeTiers,CodeAffaire,Etablissement,'',EnHT,SaisieContre);
End;

{***********A.G.L.***********************************************
Auteur  ...... : G. Merieux
Créé le ...... : 14/04/2000
Modifié le ... : 14/04/2000
Description .. : Alimentation Tob pour une échéance donnée
Mots clefs ... : ECHEANCE AFFAIRE
*****************************************************************}
Function RemplirTOBEcheance ( CodeAffaire : String;numech : Integer;TobEch : TOB ):boolean;
Var Q : TQuery ;
    Req : String;
BEGIN
Result:=True ;
if CodeAffaire='' then  TOBEch.InitValeurs(False)
else if CodeAffaire<>TOBEch.GetValue('AFA_AFFAIRE') then
   BEGIN
   // PL le 06/03/02 : INDEX 3
      // SELECT * : un seul enrgt chargé ... on laisse...
   req := 'SELECT * FROM FACTAFF WHERE AFA_TYPECHE="NOR" AND AFA_AFFAIRE="' +CodeAffaire+ '" AND AFA_NUMECHE=" '+ IntToStr(numech)+'"';
   Q:=OpenSQL(Req,True,-1,'',true) ;
   If (Not Q.EOF) then
        Result:=TOBEch.SelectDB('',Q)
        Else Result:=False;
   Ferme(Q);
   END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 17/05/2000
Modifié le ... : 17/05/2000
Description .. : Test l'existance d'échéances sur l'affaire
Mots clefs ... : AFFAIRE; ECHEANCE AFFAIRE
*****************************************************************}
Function ExisteEcheanceAffaire ( CodeAffaire : String; facture : string='') : Boolean;
var ISFacture : string;
BEGIN
isfacture := '';
Result:=False;
if CodeAffaire='' then Exit;
if facture <> '' then Isfacture := ' AND AFA_ECHEFACT="'+facture+'"';
// PL le 06/03/02 : INDEX 3
Result := ExisteSQL ('SELECT AFA_AFFAIRE FROM FACTAFF WHERE AFA_TYPECHE="NOR" AND AFA_AFFAIRE="' +CodeAffaire+ '"'+isfacture);
END;
{***********A.G.L.***********************************************
Auteur  ...... : Pierre LENORMAND
Créé le ...... : 17/05/2000
Modifié le ... : 17/05/2000
Description .. : Test l'existance d'un budget sur une affaire
Mots clefs ... : AFFAIRE; BUDGET AFFAIRE
*****************************************************************}
Function ExisteEcheanceBudget ( CodeAffaire : String) : Boolean;
BEGIN
Result:=False;
if CodeAffaire='' then Exit;
Result := ExisteSQL ('SELECT ABU_AFFAIRE FROM AFBUDGET WHERE ABU_TYPEAFBUDGET="PVT" AND ABU_AFFAIRE="' +CodeAffaire+ '"');
END;

Function Existeaffaire (CodeAffaire,CodeTiers : String) : Boolean;
Var stWhere : String;
BEGIN
Result:=False; stWhere:='';
if CodeAffaire='' then Exit;
if (CodeTiers <> '') then stWhere := ' AND AFF_TIERS="'+ CodeTiers + '"';
Result := ExisteSQL ('SELECT AFF_AFFAIRE FROM AFFAIRE WHERE AFF_AFFAIRE="' +CodeAffaire+ '"'+stWhere);
END;

Function ExisteTiers (CodeTiers : String) : Boolean;
BEGIN
Result:=False;
if CodeTiers='' then Exit;
Result := ExisteSQL ('SELECT T_TIERS FROM TIERS WHERE T_TIERS="' +CodeTiers+ '"');
END;

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 02/02/2000
Modifié le ... : 02/02/2000
Description .. : Recherche de la pièce associée à l'affaire (alimentation de l'argument CledocAffaire) 
Mots clefs ... : AFFAIRE; PIECE AFFAIRE
*****************************************************************}
Function  SelectPieceAffaire(CodeAffaire, StatutAffaire : String; Var CleDocAffaire : R_CLEDOC; Accepte : Boolean=False):Integer;
Var Q : TQuery;
    stSQL, NaturePiece: String;
Begin
  Result:=0;
  If (Not Accepte) and (CodeAffaire='') Then
      Begin
      Result:=-1; PGIBoxAf (TexteMsgPieceAffaire[1],TitreHalley); Exit;
      End;
  If (StatutAffaire='AFF') Then
     NaturePiece:= VH_GC.AFNatAffaire
{$IFDEF BTP}
  else if (StatutAffaire='INT') Then
     NaturePiece:= 'AFF'
  else if (StatutAffaire='APP') Then
     NaturePiece:= 'DAP'
{$ENDIF}
  Else
     NaturePiece:= VH_GC.AFNatProposition;
  //   
  if CodeAffaire = '' then exit;
  // PL le 06/03/02 : INDEX 4
  stSQL := 'SELECT GP_NATUREPIECEG, GP_DATEPIECE, GP_SOUCHE, GP_NUMERO From PIECE Where GP_AFFAIRE="'+CodeAffaire+'" AND GP_NATUREPIECEG="'+
                  NaturePiece + '"';
  if Accepte then stSQL := stSQL + ' AND (SELECT AFF_ETATAFFAIRE FROM AFFAIRE WHERE AFF_AFFAIRE = GP_AFFAIREDEVIS) <> "ENC"'; // Le devis peut être terminé
  Q := nil;
  try
    Q:=OpenSQL(stSQL,true);
    If Not Q.EOF Then
      Begin
      // Attention fait un fetch de tous les enreg. Autorisé sur une affaire car le nombre de pièce retourné est faible.
      While Not Q.EOF do BEGIN Inc(Result); Q.Next; END;
      If (Result =1) Then
          Begin
          Q.first;
          CleDocAffaire.NaturePiece:= NaturePiece;
          CleDocAffaire.DatePiece:= StrToDate(Q.FindField('GP_DATEPIECE').AsString);
          CleDocAffaire.Souche:= Q.FindField('GP_SOUCHE').AsString;
          CleDocAffaire.NumeroPiece:= Q.FindField('GP_NUMERO').AsInteger;
          CleDocAffaire.Indice:= 0;
          End
      Else
          Begin
  {$IFNDEF BTP}
          if Not(ctxGCAFF in V_PGI.PGIContexte) then
              PGIBoxAf(TraduitGa(TexteMsgPieceAffaire[2]), TitreHalley);
  {$ENDIF}
          End;
      End
  Else Begin Result:=0; End;

  Finally
    Ferme(Q);
   End;
End;

{***********A.G.L.***********************************************
Auteur  ...... : G.Mérieux
Créé le ...... : 02/02/2000
Modifié le ... : 18/10/2002
Description .. : Recherche de la pièce associée à l'affaire (alimentation de
Suite ........ : l'argument CledocAffaire)
Suite ........ : 
Suite ........ : meme fonction que Select Piece affaire, mais :
Suite ........ : * sans message
Suite ........ : * on ramène 2 champs de plus
Mots clefs ... : AFFAIRE; PIECE AFFAIRE
*****************************************************************}
Function SelectPieceAffaireBis(CodeAffaire, StatutAffaire: String; Var CleDocAffaire : R_CLEDOC ;var Sttiers,StRep:string ):Integer;
Var Q : TQuery;
    stSQL, NaturePiece: String;
Begin
Result:=0;
If CodeAffaire='' Then
    Begin
    Result:=-1; PGIBoxAf (TexteMsgPieceAffaire[1],TitreHalley); Exit;
    End;
If (StatutAffaire='AFF')Then
  NaturePiece:= VH_GC.AFNatAffaire
Else If (StatutAffaire='INT')Then
  NaturePiece:= 'AFF'
Else If (StatutAffaire='APP')Then
  NaturePiece:= 'DAP'
Else
  NaturePiece:= VH_GC.AFNatProposition;
Q := nil;
try
// PL le 06/03/02 : INDEX 4
stSQL := 'SELECT GP_NATUREPIECEG, GP_DATEPIECE, GP_SOUCHE, GP_NUMERO ,GP_TIERSFACTURE,GP_REPRESENTANT From PIECE Where GP_AFFAIRE="'+CodeAffaire+'" AND GP_NATUREPIECEG="'+
                NaturePiece + '"';
Q:=OpenSQL(stSQL,true);
If Not Q.EOF Then
    Begin
    // Attention fait un fetch de tous les enreg. Autorisé sur une affaire car le nombre de pièce retourné est faible.
    While Not Q.EOF do BEGIN Inc(Result); Q.Next; END;
    If (Result =1) Then
        Begin
        Q.first;
        CleDocAffaire.NaturePiece:= NaturePiece;
        CleDocAffaire.DatePiece:= StrToDate(Q.FindField('GP_DATEPIECE').AsString);
        CleDocAffaire.Souche:= Q.FindField('GP_SOUCHE').AsString;
        CleDocAffaire.NumeroPiece:= Q.FindField('GP_NUMERO').AsInteger;
        CleDocAffaire.Indice:= 0;
        Sttiers := Q.FindField('GP_TIERSFACTURE').AsString;
        StRep := Q.FindField('GP_REPRESENTANT').AsString;
        End;
    End
Else Begin Result:=0; End;

Finally
 Ferme(Q);
 End;
End;


Function GetTobPieceAffaire (CodeAff ,StatutAffaire : String; TobPiece : TOB) : Boolean;
Var CleDocAffaire: R_CLEDOC;
    iNbPieces: integer;
    Sql: string;
begin
Result := false;
if TobPiece = Nil then Exit;
if CodeAff = '' then Exit;
iNbPieces := SelectPieceAffaire(CodeAff, StatutAffaire, CleDocAffaire);
if (iNbPieces<>1) then exit;
Sql := '"'+ CleDocAffaire.NaturePiece
   + '";"'+ CleDocAffaire.Souche
   + '";"'+ inttostr(CleDocAffaire.NumeroPiece)
   + '";"'+ inttostr(CleDocAffaire.Indice) + '"';
Result := TobPiece.SelectDB(sql, Nil, False);
AddLesSupEntete (TOBPiece)
end;

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 27/08/2001
Modifié le ... : 27/08/2001
Description .. : Récup le total en PR des lignes de pièce d'une affaire 
Suite ........ : passée en argument
Mots clefs ... : AFFAIRE; PRIX DE REVIENT
*****************************************************************}
Function  RecupPRPieceAffaire ( TobPiece : TOB) : Double;
Var i : integer;
    TobDet : TOB;
begin
Result := 0;
if TobPiece = Nil then Exit;
if TobPiece.Detail.count = 0 then Exit;
For i := 0 To TobPiece.Detail.Count-1 do
   begin
   TobDet := TobPiece.Detail[i];
   if (TobDet.GetValue('GL_TYPELIGNE') = 'ART')  then
      Result := Result + (TobDet.GetValue('GL_PMRP') * TobDet.GetValue('GL_QTEFACT'));
   end;
end;


Function MajCodeAffaireSurPiece (OldAff,NewAff, NatPiece,NumPiece,Souche : string):integer;
Var req,Aff0,Aff1,Aff2,Aff3,Aff4,WherePiece : string;
begin
Result := 0;
if (OldAff = '') or (NewAff = '') or (NatPiece= '') then Exit;

// Where sur la pièce à modifier
WherePiece := '';
if (NumPiece <> '') And (Souche <> '') then
   WherePiece := ' AND GP_NUMERO='+ NumPiece + ' AND GP_SOUCHE="'+ Souche + '"';

{$IFDEF BTP}
BTPCodeAffaireDecoupe(NewAff, Aff0, Aff1, Aff2, Aff3, Aff4, taCreat, false);
{$ELSE}
CodeAffaireDecoupe(NewAff, Aff0, Aff1, Aff2, Aff3, Aff4, taCreat, false);
{$ENDIF}

// Modif de la piece
// PL le 05/03/02 : INDEX 4
req := 'UPDATE PIECE SET GP_AFFAIRE="'+ NewAff +'", GP_AFFAIRE1="'+ Aff1+ '", GP_AFFAIRE2="'+ Aff2+ '", GP_AFFAIRE3="'+ Aff3+ '", GP_AVENANT="'+ Aff4+ '" WHERE GP_AFFAIRE="'+ OldAff +'" AND GP_NATUREPIECEG="'+ NatPiece +'" ' + WherePiece;
Result :=ExecuteSQL(req);
if Result=0 then Exit;

// Modif des lignes
WherePiece := FindEtReplace(WherePiece,'GP_','GL_',True);
// PL le 05/03/02 : INDEX 4
req:= 'UPDATE LIGNE SET GL_AFFAIRE="'+ NewAff +'", GL_AFFAIRE1="'+ Aff1+ '", GL_AFFAIRE2="'+ Aff2+ '", GL_AFFAIRE3="'+ Aff3+ '", GL_AVENANT="'+ Aff4+ '" WHERE GL_AFFAIRE="'+ OldAff +'" AND GL_NATUREPIECEG="'+ NatPiece +'" ' + WherePiece;
ExecuteSQL(req);

end;

Function DetruitPieceAffaire(CodeTiers, CodeAffaire :String ; CleDocAffaire : R_CLEDOC):Boolean;
Begin
// Appel fonction générique de destruction de pièce ....
result := true;
End;

Procedure ChargeTreeViewPieceAffaire(TobPieceAffaire : TOB ; TVPiece:TTreeView ; CodeAffaire: String) ;
Var      Q : TQuery;
Begin
Q:=OPENSQL ('SELECT GP_NATUREPIECEG,GP_NUMERO,GP_DATEPIECE,GP_AFFAIRE FROM PIECE WHERE  GP_AFFAIRE="'
            +CodeAffaire+'" AND GP_NATUREPIECEG <>"'+ VH_GC.AFNatAffaire +
            '" AND GP_NATUREPIECEG <>"'+ VH_GC.AFNatProposition +'" ORDER BY GP_DATEPIECE',False,-1,'',true);

TOBPieceAffaire.LoadDetailDB('ViewPiece','','',Q,False);
Ferme(Q);
TobPieceAffaire.PutTreeView(TVPiece, Nil,'"Pièces";GP_NUMERO|" ("|GP_NATUREPIECEG|") du "|GP_DATEPIECE|', -1, -1);
TVPiece.FullExpand;
End;


Procedure AccesEntetePieceAffaire(TobPiece,TobAdresses : TOB ; Action : TActionFiche);
Begin
TheTob:=TobPiece ;
TobPiece.data:=TobAdresses;
// mcd 16/02/06 AglLanceFiche('AFF','AFCOMPLPIECE','','',ActionToString(Action)+';ENTETEAFFAIRE') ;
AFLanceFiche_ComplementPiece (ActionToString(Action)+';ENTETEAFFAIRE') ;
if TheTob<>Nil then
   BEGIN
   TOBPiece.UpdateDBTable;
   TheTob := NIL; //CEGID
   END;
End;



Function CalculMargesurAffaire (TotalHT, TotalPR : Double; QualMarge : string) : Double;
Var MC : double;
begin
MC :=0;
if (QualMarge <> 'CO') And (QualMarge <> 'MO') And (QualMarge <> 'PC') then QualMarge := 'PC';
if QualMarge='MO' then
   BEGIN
   MC:=TotalHT-TotalPR;
   END
else if QualMarge='CO' then
   BEGIN
   if TotalPR>0 then MC:=(TotalHT/TotalPR) else MC:=0 ;
   END
else if QualMarge='PC' then
   BEGIN
   if TotalHT>0 then MC:=(TotalHT-TotalPR)/TotalHT else MC:=0 ;
   END ;
Result := Arrondi(MC,V_PGI.OkDecV);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 19/06/2000
Modifié le ... : 19/06/2000
Description .. : Chargement de l'entête (en fonction de l'arg AvecEntete) et 
Suite ........ : des lignes de la pièce associée à l'affaire (renseigné dans la 
Suite ........ :  tobpiece passée).
Mots clefs ... : AFFAIRE;LIGNESAFFAIRE
*****************************************************************}
Procedure LectureLignesAffaire (TobPiece : Tob; AvecEntete : Boolean);
Var CodeAffaire : string;
    CleDocPieceAffaire : R_CLEDOC ;
    Nbpiece : Integer;
    Q : TQuery;
Begin
  if (TOBPiece = nil) then exit;
  CodeAffaire := TobPiece.GetValue('GP_AFFAIRE');
  FillChar(CleDocPieceAffaire,Sizeof(CleDocPieceAffaire),#0) ;
  {$IFNDEF BTP}
  NbPiece := SelectPieceAffaire(CodeAffaire, 'AFF', CleDocPieceAffaire);
  {$ELSE}
  NbPiece := SelectPieceAffaire(CodeAffaire, 'AFF', CleDocPieceAffaire);
 {$ENDIF}
  If (NbPiece =  0 ) Then begin PGIBoxAf (TexteMsgPieceAffaire[3],TitreHalley); Exit; end;



  If (NbPiece <> 1 ) Then Exit;

  if AvecEntete then
   BEGIN
        // SELECT * : un seul enrgt on laisse .....
     Q := nil;
     try
       Q:=OpenSQL ( 'SELECT * FROM PIECE WHERE '+WherePiece(CleDocPieceAffaire,ttdPiece,False),true,-1,'',true) ;
       If (Not Q.EOF) then TOBPiece.SelectDB('',Q);
     finally
       Ferme(Q) ;
     end;
   END;
        // SELECT * :  uniquement les lignes d'une piece... on laisse....
  Q := nil;
  try
    Q:=OpenSQL ( 'SELECT * FROM LIGNE WHERE '+WherePiece(CleDocPieceAffaire,ttdLigne,False)+' ORDER BY GL_NUMLIGNE',true,-1,'',true) ;
    TOBPiece.LoadDetailDB('LIGNE','','',Q,False,false) ;
  finally
    Ferme(Q) ;
  end;
End;


// Fonctions Maj Activite lors validation facture
// Typeaction = VAL pour validation facture et ANU pour annulation
Procedure MajGestionAffaire (TobPiece , TobPiece_ori,TOBPieceRG,TOBBasesRG,TOBAcomptes: Tob;typeaction : string; DEV : Rdevise; FinTravaux : boolean=false);
Begin
  if ((TobPiece = Nil) and (typeaction<>'ANU') )
                or (TobPiece_ori=Nil) then Exit;

  if (tobpiece <> Nil) and (tobpiece_ori <> Nil)  then
    if (TobPiece_ori.GetValue('GP_NATUREPIECEG') = 'FAC') and (TobPiece.GetValue('GP_NATUREPIECEG') = 'AVC') then
      typeaction := 'TRF';  // transformation FAC en AVOIR

  if (TobPiece <> NIL) and (TobPiece.GetValue('GP_AFFAIRE')= '') then Exit;

  if TobPiece_ori.GetValue('GP_AFFAIRE')= '' then Exit;
     // mcd 01/02/02 inutile de faire la porcédure sur piece/affaire
  if (TobPiece <> NIL) and ((TOBPIECE.GetValue('GP_NATUREPIECEG') = VH_GC.AFNatAffaire) or
  	 (TOBPIECE.GetValue('GP_NATUREPIECEG') =VH_GC.AFNatProposition) or
  	 (TOBPIECE.GetValue('GP_NATUREPIECEG') = 'BCE'))
     then Exit;

  if V_PGI.IoError=oeOk then MajActivite(TobPiece, TobPiece_ori,TypeAction) ;
  if V_PGI.IoError=oeOk then MajEcheance(TobPiece, TobPiece_ori,TypeAction) ;
  if V_PGI.IoError=oeOk then MajAffairePiece(TobPiece, TobPiece_ori,TypeAction);
(* Ancien positionnement
  {$IFDEF BTP}
  if (V_PGI.IoError=oeOk) and (not FinTravaux) then MajApresGeneration(TobPiece,TOBPieceRG,TOBBasesRG,TOBAcomptes,DEV);
  {$ENDIF}
*)
end;

{***********A.G.L.***********************************************
Auteur  ...... : G.Merieux
Créé le ...... : 06/09/2001
Modifié le ... :   /  /
Description .. : Maj de l'activite en création ou
Suite ........ : annulation de facture
Mots clefs ... : GIGA;ACTIVITE;FACTURE AUTO
*****************************************************************}
Procedure MajActivite (TobPiece , TobPiece_ori: Tob;typeaction : string);
Var Numanc, NumNouv, Req,Req1,Req2: string;
Begin
if (TOBPiece = nil) and ((TypeAction = 'VAL') or (TypeAction = 'TRF')) then exit;


NumAnc :=EncodeRefPiece(TOBPIECE_ori);
if  (TypeAction = 'VAL') or (TypeAction = 'TRF') then
Begin
  if  (TypeAction = 'TRF') then Req2 := ',ACT_ACTIVITEREPRIS = "F" ' else  Req2 := '';;
  NumNouv :=EncodeRefPiece(TOBPIECE);
  Req := 'UPDATE ACTIVITE SET ACT_NUMPIECE ='+'"'+ NumNouv+'"'
     +Req2+' WHERE ACT_NUMPIECE='+'"'+NumAnc+'"';     //GISE
End
else
Begin
  NumNouv :='';
  Req1 := '';
  if (GetParamSoc('SO_AFAPPPOINT')=true) then // gestion visa facturation sur lignes
    Req1 := ' ,ACT_ETATVISAFAC="ATT"'
          + ' ,ACT_VISEURFAC="'+V_PGI.User+'"'
          + ' ,ACT_DATEVISAFAC="'+USDateTime(idate1900)+'"';

  Req := 'UPDATE ACTIVITE SET ACT_ACTIVITEREPRIS = "F",ACT_NUMPIECE ='+'"'+ NumNouv+'"'
      + ',ACT_NUMAPPREC = 0 '
      + req1
      + ' WHERE ACT_NUMPIECE='+'"'+NumAnc+'"';
End;

if (((typeaction = 'VAL') or (TypeAction = 'TRF'))
    and (NumNouv='')) Or (NumAnc ='') then Exit;

ExecuteSQL(Req);
end;



{***********A.G.L.***********************************************
Auteur  ...... : G.Merieux
Créé le ...... : 06/09/2001
Modifié le ... :   /  /
Description .. : Maj des échéances de facturation en création ou
Suite ........ : annulation de facture
Mots clefs ... : GIGA;ECHEANCE;FACTURE AUTO
*****************************************************************}
Procedure MajEcheance (TobPiece , TobPiece_ori: Tob;typeaction : string);
Var Numanc, NumNouv, Req ,Req1 : string;
Begin
if ((TypeAction = 'VAL') or (TypeAction = 'TRF')) and (TOBPiece = nil) then exit;
NumAnc :=EncodeRefPiece(TOBPIECE_ori);

if (TypeAction = 'VAL') then
Begin
  NumNouv :=EncodeRefPiece(TOBPIECE);
  Req := 'UPDATE FACTAFF SET AFA_NUMPIECE ='+'"'+ NumNouv+'"'
     +' WHERE AFA_NUMPIECE='+'"'+NumAnc+'"';
End
else
Begin
  if (TypeAction = 'TRF') then
    NumNouv :=EncodeRefPiece(TOBPIECE)
  else
    NumNouv :='';
  Req1 := ',AFA_JUSTIFBONI="",AFA_JUSTIFBONI2="",AFA_BONIMALIQTE=0' +
  				',AFA_BONIMALI=0,AFA_BONIMALIDEV=0'+
          ',AFA_BM1FO=0,AFA_BM1FODEV=0'+
          ',AFA_BM1FR=0,AFA_BM1FRDEV=0'+
          ',AFA_BM2PRQTE=0,AFA_BM2PR=0,AFA_BM2PRDEV=0'+
          ',AFA_BM2FO=0,AFA_BM2FODEV=0'+
          ',AFA_BM2FR=0,AFA_BM2FRDEV=0'+
          ',AFA_AFACTOT=0,AFA_AFACTOTDEV=0'+
          ',AFA_AFACTURERQTE=0,AFA_AFACTURER=0,AFA_AFACTURERDEV=0'+
          ',AFA_AFACTFR=0,AFA_AFACTFRDEV=0'+
          ',AFA_AFACTFO=0,AFA_AFACTFODEV=0'+
          ',AFA_ACPTEQTE=0,AFA_ACPTEPR=0,AFA_ACPTEPRDEV=0'+
          ',AFA_ACPTEFO=0,AFA_ACPTEFODEV=0'+
          ',AFA_ACPTEFR=0,AFA_ACPTEFRDEV=0'+
          ',AFA_LIBELLE1="",AFA_LIBELLE2="" ,AFA_LIBELLE3="" ';

  Req := 'UPDATE FACTAFF SET AFA_ECHEFACT = "-",AFA_NUMPIECE ='+'"'+ NumNouv+'"'
    +',AFA_ETATVISA="",AFA_VISEUR="",AFA_DATEVISA="'+UsDateTime(idate1900)+'"'
    + Req1
    +' WHERE AFA_NUMPIECE='+'"'+NumAnc+'"';

End;

if (((typeaction = 'VAL') or (TypeAction = 'TRF')) and (NumNouv='')) Or (NumAnc ='') then Exit;

ExecuteSQL(Req);
end;


Procedure MajAffairePiece(TobPiece,TobPiece_ori : Tob;typeaction : string);
Var Req,NumNouv,NumAnc,{NumPiece,}Req1, aff: string;
    //Q : TQuery;
    //TOBAffaire : TOB;
Begin
  // fct qui met à jour le dernier n° de piece de l'affaire
  if (( TypeAction = 'VAL')or(TypeAction = 'TRF')) and (TOBPiece = nil) then exit;

  NumAnc :=EncodeRefPiece(TOBPIECE_ori);
  Aff:= TobPiece_Ori.getValue('GP_AFFAIRE'); //gm 12/03/2002

  Req1 := '';
  if (TypeAction = 'VAL') or( TypeAction = 'TRF') then
  begin
    NumNouv := EncodeRefPiece(TOBPIECE);
  end else
  begin
    NumNouv := '';
    Req1 := ',AFF_NUMSITUATION=0,AFF_DATESITUATION="'+UsDateTime(Idate1900)+'"';
  end;
  Req := 'UPDATE AFFAIRE SET AFF_NUMDERGENER ="' + NumNouv+'"'
  + Req1
  +' WHERE AFF_NUMDERGENER="'+NumAnc+'"';

  ExecuteSQL(Req);
  If ToBPiece=Nil then exit ;// mcd 26/11/02 pour eviter de planter sur ligne suivante ... si Nil
End;

Procedure MajFActEclat (TobPiece : TOB );
Var
 tobProd,tobPieceDet : tob;
 Aff: Array of string;
 wi,wj,nbaff :integer;
 trouve:boolean;
Begin
  SetLength(Aff,10);
  //nbaff:=TobPiece.getvalue('gp_numero');
  Nbaff:=0;
           // on peut avoir plusieurs affaire dans une même pièce, on stock les différentes valeurs
  for wi := 0 to TObPiece.detail.count-1 do
    begin
    TobPieceDet:=TobPiece.detail[wi];
    trouve:=false;
    If NbAff >0 then
      for wj:=0 to Nbaff-1 do
        begin
        if  TobPieceDet.getvalue('GL_AFFAIRE')=Aff[wj] then Trouve:=True;
        end;
    if (Trouve=False) then begin
         Aff[nbaff]:=TobPieceDet.getvalue('GL_AFFAIRE');
         inc(Nbaff);
         if NbAff > 9 then begin
            V_PGI.IoError:=oeUnknown ;
            // fait dans une transaction, il ne faut pas de message PGIInfoAf('La Pièce comporte plus de 10 affaires différentes','');
            Exit;
            end;
       end;
    end;
  TobProd:=Nil;
  ChargeTobActivite(TobProd,TobPiece,Aff,NbAff);
  EcrireDansAfCumul(TobProd,TobPiece,Aff,NbAff);
  TobProd.free;  TobProd:=Nil;
  Aff:=Nil;
end;

Procedure EcrireDansAfCumul (TobProd,tobPiece : TOB;Aff:array of string;Nbaff:integer );
var TObAfCumul,TobPieceDet ,TobProddet,TobAfcumulDet: TOB;
 wi,wj: integer;
 Ress,typeArt,Art,TypeRess: string;
 Coeff,totvente,Totpr,tot1,Tot2:double;
 QQ:Tquery;
 NumEclatAff : Array of Integer;
begin
  Totpr := 0; TotVente := 0;
  Coeff :=0;
  TOBAfCumul:=TOB.Create('AFCUMUL',Nil,-1) ;
  SetLength(NumEclatAff,10);
  For wi:=0 to Nbaff-1 do
      begin
                  // on regarde à tout hasard si on a pour l'affaire, à la même date, une autre facture. Cumul dans ce cas
      QQ:=OpenSQL('SELECT Max(ACU_NUMECLAT) FROM AFCUMUL WHERE ACU_TYPEAC="FAC" and ACU_DATE="'
                   +UsDateTime(TobPiece.getvalue('GP_DATEPIECE'))+'" AND ACU_TIERS="'+TobPiece.getvalue('GP_TIERS')
                   +'" AND ACU_AFFAIRE="'+Aff[wi]+'"',TRUE,-1,'',true) ;
      if Not QQ.EOF then NumEclatAff[wi]:=QQ.Fields[0].AsInteger+1
                    else NumEclatAff[wi]:=1;
     If (TobPiece.getvalue('GP_NATUREPIECEG')='FRE') and (NumEclatAff[wi] <>1)
         then begin
           // les fatcures de reprises peuvent se modifier... il faut donc détruire ce qui existe
         ExecuteSql( 'DELETE AFCUMUL WHERE ACU_TYPEAC="FAC" and ACU_DATE="'
                   +UsDateTime(TobPiece.getvalue('GP_DATEPIECE'))+'" AND ACU_TIERS="'+TobPiece.getvalue('GP_TIERS')
                   +'" AND ACU_AFFAIRE="'+Aff[wi]+'" AND ACU_NUMERO='+InttoSTr(TobPiece.getvalue('GP_NUMERO'))
                   +' AND ACU_SOUCHE="'+TobPiece.getvalue('GP_SOUCHE')+'" AND ACU_NATUREPIECEG="'+TobPiece.getvalue('GP_NATUREPIECEG')
                   +'"');
         Ferme(QQ);  // on regarde si il existe une autre facture pour la même date (sur autre souche .. dans ce cas, il ne faut pas partir de 1 pour le compteur
                     // et la clé ne comporte pas le N° de souche et n° car inutile en Cut off ...
         QQ:=OpenSQL('SELECT Max(ACU_NUMECLAT) FROM AFCUMUL WHERE ACU_TYPEAC="FAC" and ACU_DATE="'
                     +UsDateTime(TobPiece.getvalue('GP_DATEPIECE'))+'" AND ACU_TIERS="'+TobPiece.getvalue('GP_TIERS')
                     +'" AND ACU_AFFAIRE="'+Aff[wi]+'"',TRUE,-1,'',true) ;
         if Not QQ.EOF then NumEclatAff[wi]:=QQ.Fields[0].AsInteger+1
                      else NumEclatAff[wi]:=1;
        end;
      Ferme(QQ);
     end;
     // fct qui boucle sur les lignes de pièces et alimente la table
     // afcumul en fct du paramétrage et du coeff calculé dans la prod
  for wi := 0 to TObPiece.detail.count-1 do
      begin
      TobPieceDet:=TobPiece.detail[wi];
      if TobPieceDet.getvalue('GL_TYPELIGNE')<>'ART' then continue;
      If TobProd<>Nil then begin
        if GetparamSoc('So_AFFACTPARRES') = 'COD' then TobProddet:=TobProd.FIndfirst(['ACT_TYPEARTICLE','ACT_CODEARTICLE','ACT_AFFAIRE'],
                [TobPieceDet.GetValue('GL_TYPEARTICLE'),TobPieceDet.GetValue('GL_CODEARTICLE'),TobPieceDet.getvalue('GL_AFFAIRE')],True)
           else if GetparamSoc('So_AFFACTPARRES') = 'TYP' then TobProddet:=TobProd.FIndfirst(['ACT_TYPEARTICLE','ACT_AFFAIRE'],
                [TobPieceDet.GetValue('GL_TYPEARTICLE'),TobPieceDet.getvalue('GL_AFFAIRE')],true)
           else     TobProddet:=TobProd.FIndfirst(['ACT_AFFAIRE'],[TobPieceDet.getvalue('GL_AFFAIRE')],True);
         end
      Else TobProdDet:=Nil;
      ress:=''; // pour passer enseuite par l'initilaisation si la tob n'est pas vide
      if TobProddet=Nil then begin
                     //rien n'a été trouvé, il faut passer sur les codes par défaut
         TobProdDet:=tob.create('pour Ok while',Nil,-1);
         Ress:= GetParamSoc ('So_AfFactRessDefaut');
         TypeRess:='SAL';
         Totvente:=0;  TotPr:=0; Coeff:=1;
         if GetparamSoc('So_AFFACTPARRES') = 'COD' then  begin
            TypeArt:= TobPieceDet.GetValue('GL_TYPEARTICLE');
            Art:= TobPieceDet.GetValue('GL_CODEARTICLE');
            end
           else if GetparamSoc('So_AFFACTPARRES') = 'TYP' then begin
            TypeArt:= TobPieceDet.GetValue('GL_TYPEARTICLE');
            If TobPieceDet.GetValue('GL_TYPEARTICLE') = 'PRE' then Art:=GetParamSoc('SO_AfFactPresDefaut')
            else If TobPieceDet.GetValue('GL_TYPEARTICLE') = 'CTR' then  begin Art:= GetParamSoc('SO_AfFactPresDefaut');TypeArt:='PRE'; end
            else If TobPieceDet.GetValue('GL_TYPEARTICLE') = 'FRA' then  Art:= GetParamSoc('SO_AfFactFraisDefaut')
            Else If TobPieceDet.GetValue('GL_TYPEARTICLE') = 'POU' then  begin Art:= GetParamSoc('SO_AfFactFraisDefaut');TypeArt:='FRA'; end
            Else If TobPieceDet.GetValue('GL_TYPEARTICLE') = 'MAR' then  Art:= GetParamSoc('SO_AfFactFourDefaut');
            end
           else begin
            Art:= GetParamSoc('SO_AfFactPresDefaut');
            TypeArt:='PRE';
            end;
         end;
     while TobProdDet <>Nil do begin    // on alimente la table afcumul. test si enrgt existe
      if ress = '' then begin    // pour ne pas alimenter si 1er passage et zones alimenter ci dessus
           if GetparamSoc('So_AFFACTPARRES') = 'COD' then begin
              art:=TobProdDet.getvalue('ACT_CODEARTICLE');
              typeart:=TobProdDet.getvalue('ACT_TYPEARTICLE');
              end
           else if GetparamSoc('So_AFFACTPARRES') = 'TYP' then  begin
                typeart:=TobProdDet.getvalue('ACT_TYPEARTICLE');
                If TypeArt = 'PRE' then Art:=GetParamSoc('SO_AfFactPresDefaut')
                else If typeArt = 'CTR' then  begin Art:= GetParamSoc('SO_AfFactPresDefaut') ;TypeArt:='PRE'; end
                else If TypeArt = 'FRA' then  Art:= GetParamSoc('SO_AfFactFraisDefaut')
                Else If TypeArt = 'POU' then  begin Art:= GetParamSoc('SO_AfFactFraisDefaut');TypeArt:='FRA'; end
                Else If TypeArt = 'MAR' then  Art:= GetParamSoc('SO_AfFactFourDefaut');
                end
           else begin
                TypeArt:='PRE';
                Art:= GetParamSoc('SO_AfFactPresDefaut');
                end;
           Ress:=TobProdDet.getvalue('ACT_RESSOURCE');
           TypeRess:=TobProdDet.getvalue('ACT_TYPERESSOURCE');
           Coeff:=TobProdDet.getvalue('COEFF');
           Totvente:=TobProdDet.getvalue('ACT_TOTVENTE');
           TotPr:=TobProdDet.getvalue('ACT_TOTPR');
           end;
          TobAfCumulDEt:=TobAfcumul.FIndfirst(['ACU_TYPEARTICLE','ACU_CODEARTICLE','ACU_RESSOURCE','ACU_AFFAIRE']
             ,[Typeart,Art,Ress,TobPieceDet.getvalue('GL_AFFAIRE')],True);
          If (TobAfCumulDet = Nil) then begin
                            // initialisation zone pour enrgt vide
             TobAfCumulDet:= Tob.create('AFCUMUL',TobAfCumul,-1);
             For wj:=0 to NbAff-1 do
                 begin
                 if TobPieceDet.getvalue('Gl_AFFAIRE')= Aff[wj] then TobAfCumulDet.putvalue('ACU_NUMECLAT',NumEclatAff[wj]);
                 end;
             TobAfCumulDet.putvalue('ACU_CODEARTICLE',Art);
             TobAfCumulDet.putvalue('ACU_ARTICLE',CodeArticleUnique2(Art,''));
             TobAfCumulDet.putvalue('ACU_TYPEARTICLE',Typeart);
             TobAfCumulDet.putvalue('ACU_RESSOURCE',Ress);
             TobAfCumulDet.putvalue('ACU_AFFAIRE',TobPieceDEt.getvalue('GL_AFFAIRE'));
             TobAfCumulDet.putvalue('ACU_AFFAIRE1',TobPieceDEt.getvalue('GL_AFFAIRE1'));
             TobAfCumulDet.putvalue('ACU_AFFAIRE2',TobPieceDEt.getvalue('GL_AFFAIRE2'));
             TobAfCumulDet.putvalue('ACU_AFFAIRE3',TobPieceDEt.getvalue('GL_AFFAIRE3'));
             TobAfCumulDet.putvalue('ACU_AFFAIRE0',Copy(TobPieceDEt.getvalue('GL_AFFAIRE'),1,1));
             TobAfCumulDet.putvalue('ACU_AVENANT',TobPieceDEt.getvalue('GL_AVENANT'));
             TobAfCumulDet.putvalue('ACU_DATE',TobPieceDEt.getvalue('GL_DATEPIECE'));
             TobAfCumulDet.putvalue('ACU_SEMAINE',NumSemaine(TobPieceDEt.getvalue('GL_DATEPIECE')));
             TobAfCumulDet.putvalue('ACU_PERIODE',GetPeriode(TobPieceDEt.getvalue('GL_DATEPIECE')));
             TobAfCumulDet.putvalue('ACU_TIERS',TobPieceDEt.getvalue('GL_TIERS'));
             TobAfCumulDet.putvalue('ACU_TYPEAC','FAC');
             TobAfCumulDet.putvalue('ACU_TYPERESSOURCE',TypeRess);
             TobAfCumulDet.putvalue('ACU_NumPiece',EncodeRefPiece(TobPieceDet));
             TobAfCumulDet.putvalue('ACU_SOUCHE',TobPieceDEt.getvalue('GL_SOUCHE'));
             TobAfCumulDet.putvalue('ACU_INDICEG',TobPieceDEt.getvalue('GL_INDICEG'));
             TobAfCumulDet.putvalue('ACU_NUMERO',TobPieceDEt.getvalue('GL_NUMERO'));
             TobAfCumulDet.putvalue('ACU_NATUREPIECEG',TobPieceDEt.getvalue('GL_NATUREPIECEG'));
             TobAfCumulDet.putvalue('ACU_DATEMODIF', V_PGI.DateEntree);
             end;
             // cumul zone pour enrgt à mettre à jour
          // mcd 26/09/02 il faut réfléchir si alimentation .. à éclater comme le prix, sinon, faux TobAfCumulDet.putvalue('ACU_QTEFACT',TobPieceDEt.getvalue('GL_QTEFACT'));
          TobAfCumulDet.putvalue('ACU_PVPROD',TotVente);
          TobAfCumulDet.putvalue('ACU_PVPRODDEV',TotVente);
          TobAfCumulDet.putvalue('ACU_PRPROD',TotPr);
          TobAfCumulDet.putvalue('ACU_PRPRODDEV',TotPr);
          //mcd 22/11/02 pas utilisé ...it:=Tobaf.Getvalue('ACU_PVFACT')+Arrondi(Coeff* TobPieceDet.getvalue('Gl_TOTALHT'),V_PGI.OkDecV);
          TobAfCumulDet.putvalue('ACU_PVFACT',TobAfCumulDet.Getvalue('ACU_PVFACT')+
             Arrondi(Coeff* TobPieceDet.getvalue('Gl_TOTALHT'),V_PGI.OkDecV));
          TobAfCumulDet.putvalue('ACU_PVFACTDEV',TobAfCumulDet.Getvalue('ACU_PVFACTDEV')+
             Arrondi(Coeff* TobPieceDet.getvalue('Gl_TOTALHTDEV'),V_PGI.OkDecV));
          If TobProd<>Nil then Begin
            if GetparamSoc('So_AFFACTPARRES') = 'COD' then TobProddet:=TobProd.FIndNExt(['ACT_TYPEARTICLE','ACT_CODEARTICLE','ACT_AFFAIRE'],
                                                      [TobPieceDet.GetValue('GL_TYPEARTICLE'),TobPieceDet.GetValue('GL_CODEARTICLE'),TobPieceDet.GetValue('GL_AFFAIRE')],True)
               else if GetparamSoc('So_AFFACTPARRES') = 'TYP' then TobProddet:=TobProd.FIndnext(['ACT_TYPEARTICLE','ACT_AFFAIRE'],
                                                      [TobPieceDet.GetValue('GL_TYPEARTICLE'),TobPieceDet.GetValue('GL_AFFAIRE')],true)
               else     TobProddet:=TobProd.FIndNext(['ACT_AFFAIRE'],[TobPieceDet.GetValue('GL_AFFAIRE')],True);
               end
            else TobProdDet:=Nil;
          ress:='';
          end;
      end ;
      // on regarde si mtt éclatée = mtt facture origine. sinon mis eécart sur 1er
  For wj:=0 to Nbaff-1 do
    begin
    Tot1:=TOBAfCumul.Somme('ACU_PVFACT',['ACU_AFFAIRE'],[Aff[wj]],False);
    Tot2:=TOBPiece.Somme('GL_TOTALHT',['GL_TYPELIGNE','GL_AFFAIRE'],['ART',aff[wj]],False);
    If Tot1<>Arrondi(Tot2,V_PGI.OkDecV) then begin
        TobAfCumulDet:=TobAfcumul.FIndfirst(['ACU_AFFAIRE'],[Aff[wj]],True);;
        if TobAfCumuLDet <>Nil then begin
            TobAfCumulDet.putvalue('ACU_PVFACT',TobAfCumulDet.Getvalue('ACU_PVFACT')+Arrondi(Tot2,V_PGI.OkDecV)-ToT1);
           end;
       end;
    Tot1:=TOBAfCumul.Somme('ACU_PVFACTDEV',['ACU_AFFAIRE'],[Aff[wj]],False);
    Tot2:=TOBPiece.Somme('GL_TOTALHTDEV',['GL_TYPELIGNE','GL_AFFAIRE'],['ART',Aff[wj]],False);
    If Tot1<>Arrondi(Tot2,V_PGI.OkDecV) then begin
        TobAfCumulDet:=TobAfcumul.FIndfirst(['ACU_AFFAIRE'],[Aff[wj]],True);;
        if TobAfCumuLDet <>Nil then begin
            TobAfCumulDet.putvalue('ACU_PVFACTDEV',TobAfCumulDet.Getvalue('ACU_PVFACTDEV')+Arrondi(Tot2,V_PGI.OkDecV)-Tot1);
           end;
       end;
    end;
  TobAfCumul.InsertOrUpdatedb(False);
  TobAfCumul.free;
  NumEclatAff:=Nil;
end;

Procedure ChargeTobActivite (Var TobProd: TOB;tobPiece : TOB ;Aff:array of string;Nbaff:integer);
Var Req ,Req1 ,zwhere,wheredate,zorder: string;   //,date
 QQ:Tquery;
 TobDet:TOB;
 ddeb,dfin : TdateTime;
 CumTot,Coeff: Double;
 wi:integer;
iAnneeDateFin, iMoisDateFin, iJourDateFin : word;
//iMoisDateCalc, iJourDateCalc:word;
Begin
  If Nbaff =0   then exit; // mcd 02/08/01 cas ou rien dans la pièce,même aucune ligne (cas qui ne devrait pas se produire mais existe dans notre base !!!
  If (NbAff=1) and (Aff[0]='') then Exit; // cas au pièce associé à aucune affaire, on sort
  Ddeb :=Idate1900; Dfin := Idate2099;
// fct qui charge l'activité sur la période voulue et en fct de
// paraèltres de la base, afin d'avoir la abse pou fact eclat par assistant
    // sélect zones en fct option d'éclatement
  TobProd := Tob.Create('la production',NIL,-1);
  Req1 	 := 'SELECT SUM(ACT_TOTVENTE) AS ACT_TOTVENTE,SUM(ACT_TOTPR) AS ACT_TOTPR,ACT_TYPERESSOURCE,ACT_RESSOURCE,ACT_AFFAIRE';
  if GetparamSoc('So_AFFACTPARRES') = 'COD' then req1:=req1+',ACT_CODEARTICLE,ACT_TYPEARTICLE'
   else if GetparamSoc('So_AFFACTPARRES') = 'TYP' then req1:=req1+',ACT_TYPEARTICLE';
  req1:=req1 +' FROM ACTIVITE ';
    // select activité sur dates en fct paramètres
  if (GetParamSoc('SO_AFFEdatFinAct')='DFA') then
     begin      // à partir de la date de fin activité
     if (GetParamSoc('SO_TYPEDATEFINACTIVITE')='DAF') then DFin:=VH_GC.AFDateFinAct
     else if (GetParamSoc('SO_TYPEDATEFINACTIVITE')='FMC') then Dfin:=FINDEMOIS(V_PGI.DateEntree)
     else if (GetParamSoc('SO_TYPEDATEFINACTIVITE')='FM1') then
        begin
        // Fin du mois courant + 1
        DecodeDate( V_PGI.DateEntree, iAnneeDateFin, iMoisDateFin, iJourDateFin);
        iJourDateFin:=1;
        iMoisDateFin:=iMoisDateFin+1;
        if  (iMoisDateFin>12) then begin iMoisDateFin:=1; iAnneeDateFin:=iAnneeDateFin+1; end;
        DFin:=FINDEMOIS(EncodeDate(iAnneeDateFin, iMoisDateFin, iJourDateFin));
        end
     else if (GetParamSoc('SO_TYPEDATEFINACTIVITE')='FAC') then
        begin
        // Fin d'annee courante
        DecodeDate( V_PGI.DateEntree, iAnneeDateFin, iMoisDateFin, iJourDateFin);
        DFin:=FINDEMOIS(EncodeDate(iAnneeDateFin, 12, 1));
        end;
     end
  else Dfin := tobPiece.getvalue('Gp_Datepiece');  // à partir date de la pièce
    //date début cabinet
  if GetParamSoc('So_AFFACTMODEPEC') = 'DE' then ddeb :=GetparamSoc('So_AfDAteDEbCAb')
   else   if GetParamSoc('So_AFFACTMODEPEC') = 'MC' then begin
          // mois précédent la facture
         ddeb :=PlusMois(tobPiece.getvalue('Gp_Datepiece'), -1);
         ddeb := plusdate (Ddeb,1,'J');
         end
   else   if GetParamSoc('So_AFFACTMODEPEC') = 'TC' then begin
          // trimestre précédent la facture
         ddeb :=PlusMois(tobPiece.getvalue('Gp_Datepiece'), -3);
         ddeb := plusdate (Ddeb,1,'J');
         end
  else   if GetParamSoc('So_AFFACTMODEPEC') = 'DF' then begin
            // depuis date dernière facture + 1 jour
            // pris dans les pièces car l'affaire stock la dernieère pièce (éventuellement facture prov ...)
         req:= 'SELECT MAX (gp_datepiece) FROM PIECE WHERE (gp_naturepieceg="FAC" or gp_naturepieceg="AVC" or gp_naturepieceg="FRE") and GP_AFFAIRE="'+
             TobPiece.getValue('Gp_AFFAIRE')+'" and GP_DATEPIECE<"'+ UsDateTime(tobPiece.getvalue('Gp_Datepiece'))+'"';
         ddeb:=idate1900;
          QQ := nil;
          Try;
            QQ := OpenSQL(Req,true,-1,'',true) ;
            If Not QQ.EOF then begin
               ddeb:=  QQ.Fields[0].AsDateTime;
               if ddeb=0 then ddeb:=idate1900;
               if ddeb <> idate1900 then ddeb := plusdate (Ddeb,1,'J');
               end;
          Finally
          Ferme(QQ);
          End;
        end
   else   if GetParamSoc('So_AFFACTMODEPEC') = 'DA' then begin
           // depuis date dernière appréciation +1 jour
           // fait à partir des pièces et non pas des renseignement affaire pour avoir l'info précédente et non pas en cours de traitement
         req:= 'SELECT MAX (gp_datepiece) FROM PIECE WHERE (gp_naturepieceg="FAC" or gp_naturepieceg="AVC" or gp_naturepieceg="FRE") and GP_AFFAIRE="'+
             TobPiece.getValue('GP_AFFAIRE')+'" and GP_DATEPIECE<"'+ UsDateTime(tobPiece.getvalue('Gp_Datepiece'))+'" AND GP_FACREPRISE > 5000';
         ddeb:=idate1900;
          QQ := nil;
          Try;
            QQ := OpenSQL(Req,true,-1,'',true) ;
            If Not QQ.EOF then begin
               ddeb:=  QQ.Fields[0].AsDateTime;
               if ddeb=0 then ddeb:=idate1900;
               if ddeb <> idate1900 then ddeb := plusdate (Ddeb,1,'J');
               end;
          Finally
          Ferme(QQ);
          End;
         end ;

  zwhere := ' WHERE  act_typeactivite="REA" ';
  If Nbaff >0 then begin
      Zwhere:= Zwhere+' AND (';
      for wi:=0 to Nbaff-1 do
          begin
          If Wi > 0 then Zwhere:=Zwhere + ' OR ';
          Zwhere := zwhere + ' ACT_AFFAIRE="'+Aff[wi]+'"';
          end;
      Zwhere:=Zwhere + ')';
      end ;                            
  AjoutTypeArticle('ACT',Zwhere);
  wheredate := ' AND ACT_DATEACTIVITE <= "'+ usdatetime(dfin)+'"';
  wheredate := wheredate + ' AND ACT_DATEACTIVITE >= "'+ usdatetime(ddeb)+'" AND ACT_ETATVISA="VIS"';

    // group by  en fct option d'éclatement
  zorder := '  group by act_affaire,act_typeressource,act_ressource';
  if GetparamSoc('So_AFFACTPARRES') = 'COD' then zorder:=zorder+',ACT_TYPEARTICLE,ACT_CODEARTICLE'
   else if GetparamSoc('So_AFFACTPARRES') = 'TYP' then zorder:=zorder+',ACT_TYPEARTICLE';

  Req 	 := Req1+zwhere+wheredate+zorder ;

  QQ := nil;
  Try
    QQ := OpenSQL(Req,true,-1,'',true) ;
    If Not QQ.EOF then TobProd.LoadDetailDB('ACTIVITE','','',QQ,True);
  Finally
    Ferme(QQ);
  End;
     
If TobProd =Nil then exit;
       // calcul du coeff pour chaque ligne en fct paramétrage
for wi := 0 to TObProd.detail.count-1 do
  begin
    TobDet := TobProd.detail[wi];
    TobDet.AddChampSup('COEFF',False);   // ajout champ coeff su enrgt
    if GetparamSoc('So_AFFACTPARRES') = 'COD' then CumTot := TOBProd.Somme('ACT_TOTVENTE',['ACT_TYPEARTICLE','ACT_CODEARTICLE','ACT_AFFAIRE'],
                   [TobDet.GetValue('ACT_TYPEARTICLE'),TobDet.GetValue('ACT_CODEARTICLE'),TobDet.GetValue('ACT_AFFAIRE')],False)
     else if GetparamSoc('So_AFFACTPARRES') = 'TYP' then CumTot := TOBProd.Somme('ACT_TOTVENTE',['ACT_TYPEARTICLE','ACT_AFFAIRE'],
          [TobDet.GetValue('ACT_TYPEARTICLE'),TobDet.GetValue('ACT_AFFAIRE')],False)
     else     CumTot := TOBProd.Somme('ACT_TOTVENTE',['ACT_AFFAIRE'],[TobDet.GetValue('ACT_AFFAIRE')],False);
     Coeff:=0;
    if CumTot <>0 then coeff := arrondi((TObDet.GetValue('ACT_TOTVENTE') / cumtot),V_PGI.OkDecV);
    TobDet.PutValue('COEFF',Coeff)
    end;
       // On regarde si pour chaque cas, la somme des coeff est bien 1
       // sinon, arrondi sur le 1er enrgt trouvé
for wi := 0 to TObProd.detail.count-1 do
  begin
    TobDet := TobProd.detail[wi];
    if GetparamSoc('So_AFFACTPARRES') = 'COD' then CumTot := TOBProd.Somme('COEFF',['ACT_TYPEARTICLE','ACT_CODEARTICLE','ACT_AFFAIRE'],
                   [TobDet.GetValue('ACT_TYPEARTICLE'),TobDet.GetValue('ACT_CODEARTICLE'),TobDet.GetValue('ACT_AFFAIRE')],False)
     else if GetparamSoc('So_AFFACTPARRES') = 'TYP' then CumTot := TOBProd.Somme('COEFF',['ACT_TYPEARTICLE','ACT_AFFAIRE'],
          [TobDet.GetValue('ACT_TYPEARTICLE'),TobDet.GetValue('ACT_AFFAIRE')],False)
     else     CumTot := TOBProd.Somme('COEFF',['ACT_AFFAIRE'],[TobDet.GetValue('ACT_AFFAIRE')],False);
    if CumTot <>1 then begin
       coeff := 1 - CumTot;
       TobDet.PutValue('COEFF',TobDet.GetValue('COEFF')+Coeff)
       end
    else break; // mcd 25/10/02 si = 1 on arrête le test
    end;

end;


//  maj du n° piece dans l'affaire (Si saisie manuelle, ne gere pas le multi_affaire)
Procedure MajAffaire(TobPiece,TOBAcomptes : Tob; CodeAffaireAvenant, typeaction : string;Action : TActionFiche; Duplic:Boolean;SaisieAvanc:boolean; CodeEtat : String='');
Var CodeAffaire , Req,NumNouv: string;
Begin
if ((TypeAction = 'VAL') or (TypeAction = 'TRF')) and (TOBPiece = nil) then exit;
CodeAffaire := TobPiece.GetValue('GP_AFFAIRE');

{$IFDEF BTP}
// Mise à jour sous-affaire pour les devis et les etudes
if (TobPiece.Getvalue ('GP_NATUREPIECEG') = VH_GC.AFNatAffaire) or
   (TobPiece.Getvalue ('GP_NATUREPIECEG') = VH_GC.AFNatProposition) or
   (TobPiece.Getvalue ('GP_NATUREPIECEG') = 'DE') or
   (TobPiece.Getvalue ('GP_NATUREPIECEG') = 'BCE') or
   (TobPiece.Getvalue ('GP_NATUREPIECEG') = 'DAP') then
   begin
	 MajSousAffaire(TobPiece,TOBAComptes,CodeAffaireAvenant,Action,Duplic,SaisieAvanc,'',CodeEtat);
   exit;
   end;
if TobPiece.Getvalue ('GP_NATUREPIECEG') = 'AFF' then exit;
{$ENDIF}

// pas de maj de l'affaire depuis une pièce commande ou devis
if TobPiece.Getvalue ('GP_NATUREPIECEG') = VH_GC.AFNatAffaire then exit;
if TobPiece.Getvalue ('GP_NATUREPIECEG') = VH_GC.AFNatProposition then exit;
// pas de maj depuis les pièces d'achat pa le 26/07/2001
if GetInfoParPiece(TobPiece.Getvalue('GP_NATUREPIECEG'),'GPP_VENTEACHAT') <> 'VEN' then Exit;
if Codeaffaire = '' then Exit;

if (TypeAction = 'VAL') or (TypeAction = 'TRF') then
 Begin
  NumNouv := EncodeRefPiece(TOBPIECE);
  Req := 'UPDATE AFFAIRE SET AFF_NUMDERGENER ='+'"'+ NumNouv+'"'
     +' WHERE AFF_AFFAIRE='+'"'+CodeAffaire+'"';
  ExecuteSQL(Req);
 End;

End;

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 09/03/2000
Modifié le ... : 09/03/2000
Description .. : Calcul des dates d'exercices sur les missions comptables
Mots clefs ... : AFFAIRE;EXERCICE
*****************************************************************}
function CalculDateExercice(TypeExer: string; MoisCloture: integer; Exercice: string; var DebutExer, Finexer, DebutInter, FinInter, DebutFact, FinFact,
  Liquidative, Garantie, cloture: TDateTime; SurExercice: Boolean): Boolean;
var AnneeDeb, AnneeFin, MoisDeb, Moisfin, NbMois: Integer;
begin
  result := true;
  if SurExercice then
  begin

    // mcd 13/11/00 si AAAA/MM c'est année de fin
    if (MoisCloture < 1) or (MoisCloture > 12) then MoisCloture := 12;
    //InterpreteZoneExercice (TypeExer,Exercice,AnneeDeb, AnneeFin, MoisDeb,False);
    InterpreteZoneExercice(TypeExer, Exercice, AnneeDeb, AnneeFin, MoisFin, False);
    // Traitement du mois basé sur la zone exercice si renseignée sinon Mois clôture client
    {mcd 13/11/00 if (MoisDeb = 0) then BEGIN MoisDeb := MoisCloture + 1; MoisFin := Moiscloture END
                     else BEGIN Moisfin := MoisDeb -1; END; }
    if (MoisFin = 0) then
    begin
      MoisDeb := MoisCloture + 1;
      MoisFin := Moiscloture
    end
    else
    begin
      MoisDeb := MoisFin + 1;
    end;
    if MoisDeb > 12 then MoisDeb := MoisDeb - 12;
    if Moisfin < 1 then Moisfin := MoisFin + 12;
    if (TypeExer = 'A') then
    begin
      if MoisCloture = 12 then AnneeDeb := AnneeFin else AnneeDeb := AnneeFin - 1;
    end;
    if AnneeDeb <= 0 then AnneeDeb := CurrentYear;
    // mcd 13/11/00 if MoisDeb = 1 then AnneeFin:= AnneeDeb else AnneeFin := AnneeDeb+1;
    if MoisFin = 12 then AnneeDeb := AnneeFin else AnneeDeb := AnneeFin - 1;
    DebutExer := EncodeDate(AnneeDeb, MoisDeb, 01);
    FinExer := EncodeDate(AnneeFin, MoisFin, 01);
    FinExer := finDeMois(FinExer);
  end;

  // Traitement des dates d'interventions
  DebutInter := DebutExer;
  NbMois := ValeurI(GetParamSocSecur('SO_AFCALCULFIN', 0)) ;//mcd 21/09/2006 suppression Vh8gc

  if NbMois <> 0 then FinInter := PlusDate(FinExer, NbMois, 'M') else FinInter := FinExer;
  FinInter := finDeMois(FinInter);

  // Traitement des dates de facturations
  DebutFact := findeMois(DebutInter);
  FinFact := FinInter;
  // mcd 04/10/02 prise en compte paramètre
  NbMois := GetParamsocSecur('SO_AFFINFACTURAT',0);
  if NbMois <> 0 then FinFact := PlusDate(FinExer, NbMois, 'M') else FinFact := FinExer;
  FinFact := finDeMois(FinFact);

  // Traitement de la date de facture liquidative
  if (GetParamsocSecur('SO_AFGERELIQUIDE',False) = true) then
  begin
    NbMois := ValeurI(GetParamSocSecur('SO_AFCALCULLIQUID', 0)) ; //mcd 21/09/2006 suppression vh_gc

    if NbMois <> 0 then Liquidative := PlusDate(FinExer, NbMois, 'M') else Liquidative := FinFact;
    Liquidative := finDeMois(Liquidative);
  end
  else
    Liquidative := iDate1900;

  // Traitement de la date de fin de garantie
  NbMois :=   ValeurI(GetParamSocSecur('SO_AFCALCULGARANTI', 0)) ; //mcd 21/09/2006 suppression vh_gc
  if NbMois <> 0 then Garantie := PlusDate(FinExer, NbMois, 'M') else Garantie := FinInter;
  Garantie := finDeMois(Garantie);
  Cloture := Idate2099; //mcd 28/02/2005
  // Traitement de la date de clôture technique
(* mcd 28/02/2005 cette date n'est plus gérée, on en la calcul plus
  NbMois := VH_GC.AFAlimCloture;
  if NbMois <> 0 then Cloture := PlusDate(FinInter,NbMois,'M') else Cloture := FinInter;
  Cloture := finDeMois(Cloture);    *)
end;


{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 31/08/2000
Modifié le ... : 31/08/2000
Description .. : Calcul des dates d'affaires
Mots clefs ... : AFFAIRE;
*****************************************************************}

function CalculDateAffaire(DebutInter, FinInter: TDatetime; var DebutFact, FinFact, Liquidative, Garantie, Cloture: TDateTime): Boolean;
var NbMois: Integer;
begin
  result := true;
  // Traitement des dates de facturations
  DebutFact := FindeMois(DebutInter);
  FinFact := FinInter;
  // Traitement de la date de fin de garantie
  NbMois :=  ValeurI(GetParamSocSecur('SO_AFALIMGARANTI', 0)) ; //mcd 21/10/2006 suppression Vh8GC
  if NbMois <> 0 then Garantie := PlusDate(FinInter, NbMois, 'M') else Garantie := FinInter;
  Garantie := finDeMois(Garantie);
  // Traitement de la date de clôture technique
  NbMois := ValeurI(GetParamSocSecur('SO_AFALIMCLOTURE', 0)) ;//mcd 21/09/2006 suppression du Vh_gc
  if NbMois <> 0 then Cloture := PlusDate(FinInter, NbMois, 'M') else Cloture := FinInter;
  Cloture := finDeMois(Cloture);
end;


function InterpreteZoneExercice (TypeExer : string;var Exercice : String; Var  AnneeDeb,AnneeFin,Mois : integer; Obligatoire : Boolean) : Boolean;
var
    tmp : string;
BEGIN
     // MCD 13/11/00 MOis = Mois de fin (et non pas début)
Result := True;
   // mcd 18/09/01 if Not (ctxScot in V_PGI.PGIContexte) then Exit;
if TypeExer ='AUC' then Exit;
if Not(Obligatoire) and (Exercice = '') then Exit;
AnneeDeb := 0;  AnneeFin := 0; Mois := 0;
  // mcd 18/09/01 if TypeExer = 'AUC' then Exit ;
if Exercice = '' then BEGIN Result := False; Exit; END;

// Test si le nombre de caractères à été bien renseigné
Exercice := Trim (Exercice);
if (TypeExer = 'AM')And (Length(Exercice)<> 6) then BEGIN Result := False; Exit; END;
if (TypeExer = 'AA')And (Length(Exercice)<> 8)And (Length(Exercice)<> 4) then BEGIN Result := False; Exit; END;
if (TypeExer = 'A') And (Length(Exercice)<> 4) then BEGIN Result := False; Exit; END;

// interprétation des zones saisies
if TypeExer = 'AM'  then
    BEGIN
    // mcd 13/11/00 tmp := Copy(Exercice,1,4); if trim(tmp) <> '' then AnneeDeb := StrToInt(tmp ) else AnneeDeb:= 0;
    tmp := Copy(Exercice,1,4); if trim(tmp) <> '' then AnneeFin := StrToInt(tmp ) else AnneeFin:= 0;
    tmp := Copy(Exercice,5,2); if trim(tmp) <> '' then Mois := StrToInt(tmp) else mois := 0;
    END else
if TypeExer = 'AA'  then
    BEGIN
    tmp := Copy(Exercice,1,4); if trim(tmp) <> '' then AnneeDeb := StrToInt(tmp ) else AnneeDeb:= 0;
    tmp := Copy(Exercice,5,4);
    if trim(tmp)<> '' then AnneeFin := StrToInt(tmp)
                      else BEGIN AnneeFin:= AnneeDeb; Exercice :=Copy(Exercice,1,4)+Copy(Exercice,1,4); END;
                        // Année de fin d'exercice = Année de début d'exercice par défaut
    END else
if TypeExer = 'A'  then
    BEGIN
    tmp := Copy(Exercice,1,4); if trim(tmp) <> '' then AnneeFin := StrToInt(tmp ) else AnneeFin:= 0;
    END;

// En attente ...
(*if (TypeExer = 'XX') or (TypeExer = 'XX') then  // Test de l'année début sur 2 car
    BEGIN
    if (AnneeDeb < 0) Or (AnneeDeb > 99) then BEGIN Result := False; AnneeDeb := 0; END;
    if (AnneeDeb > 90) then  AnneeDeb := AnneeDeb + 1900 else AnneeDeb := AnneeDeb + 2000;
    END; **)
// Test de cohérence
if (TypeExer = 'AM') then  // Test du Mois
    BEGIN
    if (Mois < 1) Or (Mois > 12) then BEGIN Result := False; Mois := 0; END;
    if (Mois <> 1) then AnneeDeb := AnneeFin - 1 else AnneeDeb := AnneeFin;
    END;
if (TypeExer = 'AA') then  // Test comparatif des 2 années
    BEGIN
    if (AnneeDeb > AnneeFin) then BEGIN AnneeFin := AnneeDeb; Result := false; END;
    END;

END;

function FormatZoneExercice (TypeExer : string; bTiretStocke : Boolean = False) : String;
var tiret : string;
BEGIN
Result := '';
if bTiretStocke then tiret :='1' else tiret :='0';
if TypeExer = 'AUC' then Exit else
if TypeExer = 'A' then Exit else
if TypeExer = 'AM'  then Result := '!9999\-00;'+tiret+';_' else
if TypeExer = 'AA' then Result := '!9999\-0000;'+tiret+';_'
;
END;


{$IFNDEF EAGLCLIENT}
Procedure FormatExerciceGrid ( Gr : THDBGrid ) ;
Var i : integer;
BEGIN
     // mcd 18/09/01 if Not(ctxScot in V_PGI.PGIContexte) then exit;
if GetParamSoc('SO_AfFormatExer')='AUC' then exit;
For i:=0 to Gr.Columns.Count-1 do
    BEGIN
    if (Gr.Columns[i].FieldName = 'AFF_AFFAIRE2') then
        Gr.columns[i].Field.EditMask := FormatZoneExercice(VH_GC.AFFORMATEXER,false);
    END;
END;
{$ENDIF}


Function GetChampsAffaire (codeAff, NomChamp : String) : String;
Var Q : TQuery ;
    TOBLocal : TOB;
BEGIN

  Result := '';

  if CodeAff='' then Exit ;

  if not(ctxAffaire in V_PGI.PGIContexte) and not(ctxGCAFF in V_PGI.PGIContexte) then Exit;

  // OPTIMISATIONS LS
  TOBlocal := TOBlesAffaires.findfirst(['AFF_AFFAIRE'],[CodeAff],true);
  if TOBLocal = nil then
  begin
    Try
      Q:= OPENSQL ('SELECT '+NomChamp + ' From AFFAIRE WHERE AFF_AFFAIRE="' + CodeAff +'"',True,-1,'',true);
      if Not Q.EOF then Result := Q.Findfield(NomChamp).AsString;
    Finally
      Ferme(Q);
    END;
  end else
  begin
    result := TOBlocal.GetString(Nomchamp);
  end;
END;

//FV1 : 10/06/2014 - FS#921 - DELABOUDINIERE : Revoir les contrôles sur appels et contrats en fonction du code état
function ControleAffaire(CodeAff, TitleMsg, TypeSaisie : String): Boolean;
Var TypeAff : String;
    CodeEtat: String;
begin

    TypeAff := copy(UpperCase (CodeAff),1,1);

    Result := True;

    if CodeAff = '' then
    begin
      Result := False;
      Exit;
    end;

    CodeEtat := GetChampsAffaire (CodeAff,'AFF_ETATAFFAIRE');

    If TypeAff = 'A' then
    begin
      Result := ValidationCodeEtatChantier(TypeSaisie, CodeEtat, TitleMsg);
    end
    else if TypeAff = 'W' then
    begin
      Result := ValidationCodeEtatAppel(TypeSaisie, CodeEtat, TitleMsg);
    end
    else If TypeAff = 'I' then
    begin
      Result := ValidationCodeEtatContrat(TypeSaisie, CodeEtat, TitleMsg);
    end;

end;

Function ValidationCodeEtatChantier(TypeSaisie, CodeEtat, TitleMsg : String) : Boolean;
begin

  Result := true;

  if CodeEtat = 'REF' then
  begin
    PGIBox (TraduireMemoire('Saisie non autorisée. Chantier refusé'),TitleMsg);
    Result := False;
  end else if CodeEtat = 'TER' then
  begin
    PGIBox (TraduireMemoire('Saisie non autorisée. Chantier Terminé'),TitleMsg);
    Result := False;
  end else if CodeEtat <> 'ACP' then
  begin
    if TypeSaisie = 'SCO' then //SCO : Saisie conso
    begin
      if (PGIAsk(TraduireMemoire('Attention : Chantier non Accepté, voulez-vous continuer ?'),TitleMsg) = mrNo) then
      begin
        result := False;
      end;
    end
    //FV1 - 10/01/2018 - FS#2844 - TEAM - Autoriser la saisie des contre-études sur devis non accepté
    else if TypeSaisie = 'BCE' then
    begin
      if not GetParamSocSecur('SO_GENCESURDEVNACPT', False) then
      begin
        PGIBox (TraduireMemoire('Saisie non autorisée. Ce Chantier n''est pas accepté.'),TitleMsg);
        Result := False;
      end;
    end
    else
    begin
      if GetParamSocSecur('SO_BTINTERDIREACHATS',False) then
      begin
        PGIBox (TraduireMemoire('Saisie non autorisée. Ce Chantier n''est pas accepté.'),TitleMsg);
        Result := False;
      end
      else
      begin
        PgiBox (TraduireMemoire('Attention, ce chantier n''est pas accepté'),TitleMsg);
      end;
    end;
  end;
end;

function ValidationCodeEtatAppel(TypeSaisie, CodeEtat, TitleMsg : String) : boolean;
begin

  Result := True;
  if CodeEtat = 'CL1' then
  begin
    PGIBox (TraduireMemoire('Saisie non autorisée. Appel clôturé'),TitleMsg);
    Result := False;
  end else if CodeEtat = 'ANN' then
  begin
    PGIBox (TraduireMemoire('Saisie non autorisée. Appel Annulé'),TitleMsg);
    Result := False;
  end else if CodeEtat = 'FIN' then
  begin
    PGIBox (TraduireMemoire('Saisie non autorisée. Appel Terminé'),TitleMsg);
    Result := False;
  end;
end;

Function ValidationCodeEtatContrat(TypeSaisie, CodeEtat, TitleMsg : String): boolean;
begin

  Result := True;

  if CodeEtat = 'REF' then
  begin
    PGIBox (TraduireMemoire('Saisie non autorisée. Contrat refusé'),TitleMsg);
    Result := False;
  end else if CodeEtat = 'TER' then
  begin
    PGIBox (TraduireMemoire('Saisie non autorisée. Contrat Terminé'),TitleMsg);
    Result := False;
  end;
end;


Function GetMoisClotureTiers (CodeTiers : String) : Integer;
Var Q:TQuery;
BEGIN
 Result := 0;
Q := nil;
Try
 Q:=OPENSQL('SELECT T_MOISCLOTURE FROM TIERS WHERE T_TIERS="'+CodeTiers+'"',True,-1,'',true);
 If Not Q.EOF Then Result := Q.Fields[0].AsInteger;
finally
 Ferme(Q);
 end;
END;

Function FabriqueWhereNatureAuxiAff ( Statut : String) : String;
Var NatPiece,NatTiers : string;
begin
Result := 'T_NATUREAUXI="CLI"';
if ctxScot in V_PGI.PGIContexte then Exit;
If (Statut='PRO')   then NatPiece:= VH_GC.AFNatProposition else NatPiece:=VH_GC.AFNatAffaire;
if (NatPiece <> '') then NatTiers :=GetInfoParPiece(NatPiece,'GPP_NATURETIERS');
if (NatTiers <> '') then Result :=FabricWhereToken(NatTiers,'T_NATUREAUXI');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 16/08/2000
Modifié le ... : 16/08/2000
Description .. : Retourne si le tiers passé est fermé ou non
Mots clefs ... : TIERS;FERME;
*****************************************************************}
Function IsTiersFerme (CodeTiers : String): Boolean;
Var Q : TQuery;
BEGIN
Result := False;
if CodeTiers = '' then Exit;
Q := OpenSQL ('SELECT T_FERME From TIERS WHERE T_TIERS="'+CodeTiers+'"',True,-1,'',true);
if Not Q.EOF then Result := (Q.Findfield('T_FERME').AsString = 'X');
Ferme(Q);
END;


Procedure AfficheTotauxDansLibelle (Sender: TObject; Ecran : TOBJECT );
Var Nam : String ;
    CC  : THLabel ;
Begin
if Sender=Nil then Exit ;
Nam:=THNumEdit(Sender).Name ; Nam:='L'+Nam ;
CC:=THLabel(TForm(Ecran).FindComponent(Nam));
if CC<>Nil then CC.Caption:=THNumEdit(Sender).Text ;
End;

(*
function GetEtatAppelsNonTermine : string;
var QQ : TQuery;
		TOBT : TOB;
		first : boolean;
    Indice : integer;
begin
	first := true;
	QQ := OpenSql ('SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE="AET" AND ISNUMERIC(CC_LIBRE)=1 AND (NOT CC_CODE IN ("ANN","CL1","FIN","TER"))',True,-1,'',true);
  if not QQ.eof then
  begin
  	TOBT := TOB.Create ('LES CODES',nil,-1);
    TOBT.LoadDetailDB('CHOIXCOD','','',QQ,false);
    for Indice := 0 to TOBT.detail.count -1 do
    begin
    	if first then
      begin
      	result := TOBT.detail[Indice].GetValue('CC_CODE');
        first := false;
      end else
      begin
      	result := result + ';'+TOBT.detail[Indice].GetValue('CC_CODE');
      end;
    end;
    TOBT.free;
  end;
  ferme(QQ);
end;
*)

function GetPlusEtatAffaire(bAffaire : Boolean) : String;
BEGIN
Result := '';
{$IFDEF BTP}
(*if baffaire then Result := ' AND (CC_LIBRE="" OR CC_LIBRE="AFF" OR CC_LIBRE="BTP" )'
            else Result := ' AND (CC_LIBRE="" OR CC_LIBRE="PRO" )' *)
if baffaire then Result := ' AND (ISNUMERIC(CC_LIBRE)=1 OR CC_LIBRE="AFF")'
            else Result := ' AND (ISNUMERIC(CC_LIBRE)=1 OR CC_LIBRE="PRO" )'

{$ELSE}
if baffaire then Result := ' AND (CC_LIBRE="" OR CC_LIBRE="AFF")'
            else Result := ' AND (CC_LIBRE="" OR CC_LIBRE="PRO")';
{$ENDIF}

END;


Function StatutReduitToComplet (stStatut : String) : String;
BEGIN
Result := 'AFF';
if(stStatut = 'P') then Result := 'PRO';
if(stStatut = 'I') then Result := 'INT';
END;

Function StatutCompletToReduit (stStatut : String) : String;
BEGIN
Result := 'A';
if(stStatut = 'PRO') then Result := 'P';
if(stStatut = 'INT') then Result := 'I';
END;

{***********A.G.L.***********************************************
Auteur  ...... : G.Merieux
Créé le ...... : 10/07/2000
Modifié le ... :   /  /
Description .. : Analyse ACTIVITEREPRISE et transformation en Combo
Suite ........ : TYPEARTICLE
Mots clefs ... : ACTIVITEREPRIS
*****************************************************************}
//zrep : Valeur de Activite Reprise (valeur de la tablette AFTACTIVITEREPRISE)
//zcombo : valeur correspondante pour alimenter le multi-combo TYPEARTICLE
procedure Trt_Activiterepris(zrep : string; var zcombo : string;var toppre,topfou,topfra : boolean);
var critere,zone : string;
BEGIN
toppre:=false; topfra:=false ; topfou:=false;


if (zrep = 'ART') then zcombo := 'MAR;';
if (zrep = 'ARF') then zcombo := 'MAR;FRA;';
if (zrep = 'FRA') then zcombo := 'FRA;';
if (zrep = 'TOU') then zcombo := '';
if (zrep = 'PRE') then zcombo := 'PRE;';
if (zrep = 'ARP') then zcombo := 'PRE;MAR;';
if (zrep = 'FRP') then zcombo := 'FRA;PRE;';
if ((zrep = 'NON') or (zrep = '')) then
Begin
  zcombo := 'Aucun';
  exit;
End;

zone := zcombo;
if (zcombo = '') then
  Begin
    toppre:=true; topfra:=true ; topfou:=true;
  End;
// Analyse de la zone
Critere:=(Trim(ReadTokenSt(zone)));
While (Critere <>'') do
    BEGIN
    if Critere='PRE' then  toppre := true;
    if Critere='FRA' then  topfra := true;
    if Critere='MAR' then  topfou := true;

    Critere:=(Trim(ReadTokenSt(zone)));
    END;
END;

Function ToutSeulAff () : Boolean;
begin
     // cette fct test si des blocages affaires sont en cours.
     // a mettre à jour au fu et à mesure de nouveaux blocages AFFAIRE
result:= blocage (['Afftoutseul','NrGener'],true,'afftoutseul');
if (result = False) then begin
   result:= EstBloque ('Affaire',True); // on regarde si quelqu'un d'autre dans appli
   if result = true then begin
      bloqueur ('AffToutseul',false);
      PgiInfo ('Option interdite, un autre utilisateur utilise l''application.',TitreHalley);
      end;
   end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 21/03/2001
Modifié le ... :
Description .. : Fonction BlocageAffaire recherche dans la table BLOCAGEAFFAIRE la liste des blocages et alertes liés à
l'événement courant, pour le groupe de user courant. Pour tous les blocages et alertes trouvés, effectue le test décrit
dans l'enregistrement de la table (par rapport à l'état de l'affaire ou par rapport à la date courante et les dates de l'affaire)
et envoie un message en conséquence. Renvoie True s'il y a blocage. On peut rendre muet les blocages (pas les alertes) : AcbParlant=false,
on peut arrêter les messages dès le premier blocage trouvé : AcbTous=false.
Entrées :
AcsEvtCourant : Evenement duquel on appel la fonction
                MAF	Modification Affaire
                SAT	Saisie activité
                SAH	Saisie achats
                FAC	Facturation
AcsAffaireCourante : Code affaire en cours
AcsGroupeUser : Groupe de user en question (souvent en cours : V_PGI.groupe)
AcdDateCourante : Date courante pour comparer avec les dates de l'affaire
AcbParlant : true = blocages avec messages, false = blocages sans message (ne concerne pas les alertes)
AcbTous : true = enchaîne les blocages, false = sort au premier blocage
AtobAffaires : pas obligatoire, liste des affaires deja lues

Sorties :
AvdDebutItv : Date de début d'intervalle OK
AvdFinItv : Date de Fin d'intervalle OK
Si AvdDebutItv = idate1900 et AvdFinItv = idate2099, le blocage est sur l'état

Mots clefs ... : BLOCAGEAFFAIRE
*****************************************************************}
function T_OP.DonneesOperation(AcsChampFormule:Hstring):variant;
begin

        Result :=0;
        AcsChampFormule:=AnsiUppercase(AcsChampFormule);
        if (AcsChampFormule= 'DATECOURANTE') then
            result := DateCourante
        else
        if (AcsChampFormule= 'DATEAFFAIRE') then
            result := tobAffaire.GetValue(rechdom('AFTYPEBLOCAGEAFFAIRE', TypeBlocage, true))
        else
        ;
end;

Function BlocageAffaire( AcsEvtCourant, AcsAffaireCourante, AcsGroupeUser : string; AcdDateCourante : TDateTime;
                        AcbBlocParlant,AcbAlerteParlant,AcbTous : boolean;
                        var AvdDebutItv, AvdFinItv :TDateTime; AtobAffaires:TOB) : T_TypeBlocAff;
Var
Q : TQuery ;
tobBlocage, tobAffaire:TOB;
i:integer;
req : string;
Operation: Hstring;
op:T_OP;
vResultat:variant;
bMessOK:boolean;
TitreMess, Mess, vargroupeConf, sListeGroupe:string;
dDateInter:TDateTime;

BEGIN

  op:=nil; tobBlocage:=nil;
  AvdDebutItv:=idate1900;
  AvdFinItv:=idate2099;
  bMessOK:=False;
  Result:=tbaAucun;
  if (AcsEvtCourant='') or (AcsAffaireCourante='') then exit;
  if (AcdDateCourante<iDate1900) or (AcdDateCourante>iDate2099) then exit;
  tobAffaire:=TOB.Create('AFFAIRE',Nil,-1) ;
  try
    if Not AFRemplirTOBAffaire( AcsAffaireCourante, tobAffaire, AtobAffaires) then exit;

    //////////////////////
    // Partie de gestion des blocage dus à la confidentialité
    //////////////////////
    if (tobAffaire<>nil) and (GetParamsoc('SO_AFSAISAFFINTERDIT')=true) and Not (V_PGI.Superviseur) and Not (V_PGI.Controleur) then
        // Si la saisie sur les affaires non affectées est choisie dans les paramètres
        // et que le user n'est ni superviseur, ni controleur, on doit controler le blocage
        begin
        TitreMess := 'Confidentialité sur l''affaire : ';
        if (GetParamsoc('SO_AFTYPECONF')='AGR') and (tobAffaire.GetValue('AFF_GROUPECONF')<>'') then
            // Confidentialité par groupe de conf
            begin
            TitreMess := TitreMess + 'par groupe';
            vargroupeConf := VH_GC.AfGereCritGroupeConf;
            sListeGroupe := ReadTokenPipe(vargroupeConf,'ComboPlus;');
            if (pos(tobAffaire.GetValue('AFF_GROUPECONF')+';',sListeGroupe)=0) then
                begin
                Mess := TraduitGA('L''utilisateur ne fait pas partie du groupe de travail de l''affaire.' );
                Result:=tbaConfident;
                end;
            end
        else
        if (GetParamsoc('SO_AFTYPECONF')='A01') then
            // Confidentialité par responsable
            begin
            TitreMess := TitreMess + 'par responsable';
            if (tobAffaire.GetValue('AFF_RESPONSABLE')<>VH_GC.RessourceUser) then
                begin
                Mess := TraduitGA('L''utilisateur n''est pas le responsable de l''affaire.' );
                Result:=tbaConfident;
                end;
            end
        else
        if (GetParamsoc('SO_AFTYPECONF')='AS1') then
            // Confidentialité par assistant libre 1
            begin
            TitreMess := TitreMess + 'par ressource libre 1';
            if (tobAffaire.GetValue('AFF_RESSOURCE1')<>VH_GC.RessourceUser) then
                begin
                Mess := TraduitGA('L''utilisateur n''est pas une ressource libre de l''affaire.' );
                Result:=tbaConfident;
                end;
            end
        else
        if (GetParamsoc('SO_AFTYPECONF')='AS2') then
            // Confidentialité par assistant libre 2
            begin
            TitreMess := TitreMess + 'par ressource libre 2';
            if (tobAffaire.GetValue('AFF_RESSOURCE2')<>VH_GC.RessourceUser) then
                begin
                Mess := TraduitGA('L''utilisateur n''est pas une ressource libre de l''affaire.' );
                Result:=tbaConfident;
                end;
            end
        else
        if (GetParamsoc('SO_AFTYPECONF')='AS3') then
            // Confidentialité par assistant libre 3
            begin
            TitreMess := TitreMess + 'par ressource libre 3';
            if (tobAffaire.GetValue('AFF_RESSOURCE3')<>VH_GC.RessourceUser) then
                begin
                Mess := TraduitGA('L''utilisateur n''est pas une ressource libre de l''affaire.' );
                Result:=tbaConfident;
                end;
            end;

        // Exploitation des résultats du blocage
        if AcbBlocParlant and (Result<>tbaAucun) then
            PGIInfoAf(Mess, TraduitGA(TitreMess));

        if (Result<>tbaAucun) then exit;
        end;

    //////////////////////
    // Partie de gestion des blocages de l'affaire
    //////////////////////
    op := T_OP.Create;
    op.TobAffaire := tobAffaire;
    op.DateCourante := AcdDateCourante;

    tobBlocage := TOB.Create ('Les Blocages', Nil, -1);
    // SELECT * : on a vraiment besoin de tous les éléments (nombre restreint)
    req := 'SELECT * FROM BLOCAGEAFFAIRE WHERE ABA_EVENEMENTAFF="'+ AcsEvtCourant
            +'" AND (ABA_GROUPE="' + AcsGroupeUser + '" OR ABA_GROUPE="")';

    Q := nil;
    try
      Q:=OpenSQL(Req,True,-1,'',true) ;
      If (Not Q.EOF) then
        tobBlocage.LoadDetailDB('BLOCAGEAFFAIRE','','',Q,True);
    finally
      Ferme(Q);
    end;

    if tobBlocage <> Nil then
        begin
        for i:=0 to tobBlocage.Detail.Count - 1 do
            begin
            TitreMess:='';
            Mess:='';
            // Blocage sur l'etat
            if (tobBlocage.detail[i].getValue('ABA_TYPEBLOCAGE') = 'EAF') then
                begin
                if (tobBlocage.detail[i].getValue('ABA_ETATAFFAIRE')=tobAffaire.GetValue('AFF_ETATAFFAIRE')) then
                    begin
                    Mess := TraduitGA('l''affaire est '+ rechdom('AFETATAFFAIRE',tobAffaire.GetValue('AFF_ETATAFFAIRE'),false));
                    Result:=tbaEtat;
                    end;
                end
            else
                begin
                Operation := '{SI([DATECOURANTE]'
                            + rechdom('AFOPERATEUR',tobBlocage.detail[i].getValue('ABA_OPERATEUR'), true)
                            + '[DATEAFFAIRE];true;false)}';

                op.TypeBlocage := tobBlocage.detail[i].getValue('ABA_TYPEBLOCAGE');
                vResultat := GFormule(Operation, op.DonneesOperation, nil, 1);


                // On remplit les dates de l'intervalle OK en cas de blocage uniquement
                if (tobBlocage.detail[i].getValue('ABA_ALERTE')<>'X') then
                    begin
                    dDateInter := TDateTime(tobAffaire.GetValue(rechdom('AFTYPEBLOCAGEAFFAIRE', tobBlocage.detail[i].getValue('ABA_TYPEBLOCAGE'), true)));
                    if (tobBlocage.detail[i].getValue('ABA_OPERATEUR')='IEG') then
                        if (AvdDebutItv < dDateInter+1) then AvdDebutItv := dDateInter+1;
                    if (tobBlocage.detail[i].getValue('ABA_OPERATEUR')='INF') then
                        if (AvdDebutItv < dDateInter) then AvdDebutItv := dDateInter;
                    if (tobBlocage.detail[i].getValue('ABA_OPERATEUR')='SEG') then
                        if (AvdFinItv > dDateInter-1) then AvdFinItv := dDateInter-1;
                    if (tobBlocage.detail[i].getValue('ABA_OPERATEUR')='SUP') then
                        if (AvdFinItv > dDateInter) then AvdFinItv := dDateInter;
                    end;

                // Message résultat
                if (vResultat <> 'Erreur') and (vResultat <> '') then
                if boolean(vResultat) then
                    begin
                    Mess := TraduitGA('la date saisie est '
                            + rechdom('AFOPERATEUR',tobBlocage.detail[i].getValue('ABA_OPERATEUR'),false))
                            + ' au champ "'
                            + rechdom('AFTYPEBLOCAGEAFFAIRE',tobBlocage.detail[i].getValue('ABA_TYPEBLOCAGE'),false)
                            + '" de l''affaire';
                    Result:=tbaDates;
                    end;
                end;

                if (Mess<>'') then
                    begin
                    if (tobBlocage.detail[i].getValue('ABA_ALERTE')='X') then
                        begin
                        TitreMess := TraduitGA('Alerte sur l''affaire');
                        Mess := 'Attention, ' + Mess;
                        Result := tbaAucun; // PL le 14/11/02 : en alerte on ne bloque pas !!!
                        if AcbAlerteParlant then bMessOK:=true;
                        end
                    else
                        begin
                        TitreMess := TraduitGA('Blocage sur l''affaire');
                        Mess := 'Blocage : ' + Mess;
                        // AcbBlocParlant,AcbAlerteParlant
                        if AcbBlocParlant then bMessOK:= true;
                        end;

                   if bMessOK then PGIInfoAf(Mess, TitreMess);
                    if (Result<>tbaAucun) and Not AcbTous then exit;
                    end;
            end;
        end;

  finally
    tobAffaire.Free;
    tobBlocage.Free;
    op.Free;
  end;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 07/11/2002
Modifié le ... :

cette fonction rend une TOB contenant le champ ABA_ETATAFFAIRE
de tous les états bloqués

Mots clefs ... : BLOCAGEAFFAIRE; ETAT;
*****************************************************************}
function TOBEtatsBloquesAffaire (AcsEvtCourant, AcsAlerteON, AcsGroupeUser : string) : TOB;
var
  req : string;
  tobBlocage : TOB;
  QQ : TQuery;
begin
  //Result:=nil;

  tobBlocage := TOB.Create ('Blocages etats', Nil, -1);
  req := 'SELECT DISTINCT ABA_ETATAFFAIRE FROM BLOCAGEAFFAIRE WHERE ABA_EVENEMENTAFF="'+ AcsEvtCourant
            +'" AND ABA_TYPEBLOCAGE="EAF"'
            +' AND ABA_ALERTE="' + AcsAlerteON
            +'" AND (ABA_GROUPE="' + AcsGroupeUser
            + '" OR ABA_GROUPE="")';

  QQ := nil;
  try
    QQ := OpenSQL (Req, True,-1,'',true);
    If (Not QQ.EOF) then
      tobBlocage.LoadDetailDB ('BLOCAGEAFFAIRE', '', '', QQ, True)
    else
      begin
        tobBlocage.Free;
        tobBlocage:=nil;
      end;

  finally
    Ferme (QQ);
  end;

  Result := tobBlocage;

end;

function RecupLibelleTiers(code : string): string;
Var Libelle:string;
Q:TQuery;
begin
Q:=OpenSQL('SELECT T_LIBELLE FROM TIERS WHERE T_TIERS="'+code+'"',True,-1,'',true) ;
if not q.eof then
	Libelle:=Q.Findfield('T_LIBELLE').AsString
else
	Libelle:= 'Client inconnu';
Ferme(Q) ;
result := Libelle;
end;

function ActionAffaireString( AcbModifTOMAutorisee:boolean;  TypeAction:TActionFiche ):string;
begin
    if (Not AcbModifTOMAutorisee) or (TypeAction=taConsult) then
        Result:='ACTION=CONSULTATION'
    else
        Result:=ActionToString(TypeAction);
end;

Function AFRemplirTOBAffaire ( CodeAffaire : String ; TOBAffaire, TOBAffaires : TOB ) : boolean ;
Var Q : TQuery ;
T:TOB;
sPart1:string;
BEGIN
Result:=true ;
if CodeAffaire='' then
   BEGIN
   TOBAffaire.ClearDetail ;
   TOBAffaire.InitValeurs(False) ;
   END
else
   if (CodeAffaire<>TOBAffaire.GetValue('AFF_AFFAIRE')) then
      begin
      Result:=false ;
      // Si la partie 1 du code affaire est vide, ce n'est pas la peine d'aller regarder s'il existe
      sPart1 := copy(CodeAffaire,2,VH_GC.CleAffaire.Co1Lng);
      if (trim(sPart1)='') then exit;

      T:=nil;
      if (TOBAffaires<>nil) then
            T:=TOBAffaires.FindFirst(['AFF_AFFAIRE'], [CodeAffaire], false);

      if (T<>nil) then
            begin
            TOBAffaire.ClearDetail ;
            TOBAffaire.InitValeurs(False) ;
            TOBAffaire.dupliquer(T, true, true, false);
            Result:=true;
            end
      else
        if ExisteSQL('SELECT AFF_AFFAIRE FROM AFFAIRE WHERE AFF_AFFAIRE="' +CodeAffaire+ '"') then
          begin
               // SELECT * : un seul enrgt, on laisse ...
          Q := nil;
          try
            Q:=OpenSQL('SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE="' +CodeAffaire+ '"',True,-1,'',true) ;
            if Not Q.EOF then
             begin
             TOBAffaire.SelectDB('',Q);
             if (TOBAffaires<>nil) then
                begin
                T:=TOB.Create('',TOBAffaires,-1) ;
                T.dupliquer(TOBAffaire, true, true, false);
                end;
             Result:=true;
             end;
          finally
            Ferme(Q) ;
          end;
        end;          
      end;
END ;

procedure   AjoutTypeArticle(racine:string;var Zwhere:string);
Var text,critere:string;
ii:integer;
begin
// fct qui permet de ne traiter que les type article traité en fct GI/GA
//devrait être inutile chez les clients, mais dans nos base ou on a pe CTR dans des fatcures
// alors que la GI ne le traite pas, il est impossible de se comparer avec les impressions des états
Zwhere:= zwhere+' AND (';
Text:=PlusTYpeArticleText;
ii:=0;
Repeat
 Critere:=AnsiUppercase(Trim(ReadTokenSt(Text))) ;
 if Critere<>'' then
    begin
    if ii > 0 then zwhere:=zwhere+' OR ';
    zwhere:=Zwhere + racine+'_TYPEARTICLE="'+critere+'" ';
    Inc(ii);
    end;
until  Critere='';

Zwhere:=  Zwhere+') ';
end;


{***********A.G.L.***********************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... :
Modifié le ... : 02/10/2002  
Description .. : fct qui peremt d'incrémenter ou de cdécrémenter la zone
Suite ........ : exercice de GI.
Suite ........ : Fait en fct du paramétrage du champ exercice
Suite ........ : 
Suite ........ : 02/10/2002 repris du source AffaireDupli qui ne faisait que 
Suite ........ : l'incrémente et ouvert pour faire aussi décrement
Mots clefs ... : EXERCICE;N+1;N-1
*****************************************************************}
function CodificationNewExercice (Exercice : string;Plus:Boolean): string ;
Var
   annee,annee1 : Integer;
begin
 // mcd 02/10/02 : plus = vrai on ajoute, sinon on retranche (rech exer précédent)
// incrémente l'exercice de la clé
if (VH_GC.AFFormatExer = 'AM') then  // ***** Année mois *****
   begin
   annee := StrtoInt (Copy(Exercice,1,4));
   annee1 := StrtoInt(Copy(Exercice,5,2));
   if Plus then inc (annee)
    else Dec(annee);
   Result:=Format('%4.4d%2.2d',[annee,annee1]);
   end
else
if (VH_GC.AFFormatExer = 'A') then  // ***** Année ******
   begin
   annee := StrtoInt (Copy(Exercice,1,4));
   if Plus then inc (annee)
   else Dec(Annee);
   Result:=Format('%04d',[annee]);
   end
else
if (VH_GC.AFFormatExer = 'AA') then   // **** année - année *****
   begin
   annee := StrtoInt (Copy(Exercice,1,4));
   annee1 := StrtoInt (Copy(Exercice,5,4));
   if PLus then begin inc (annee); inc (annee1);  end
   else begin dec (annee); dec (annee1);  end;
   Result:=Format('%04d%4d',[annee,annee1]);
   end;
end;

Function AFMajProspectClient (CodeTiers :string):Boolean;
var StAuxi : string; 
begin
   Result := True;
   If existeSql ('SELECT T_TIERS FROM TIERS where T_NATUREAUXI="CLI" AND T_TIERS="'+CodeTiers+'"') then exit; //mcd 02/03/05 inutile si déjà client
   ExecuteSQL ('UPDATE TIERS SET T_NATUREAUXI="CLI", T_DATEPROCLI ="'+UsDateTime(V_PGI.DateEntree)+
   '" WHERE T_NATUREAUXI="PRO" AND T_TIERS="'+CodeTiers+'"');
   StAuxi :=TiersAuxiliaire (CodeTiers,False);
   ExecuteSQL ('UPDATE CONTACT SET C_NATUREAUXI="CLI" WHERE C_TYPECONTACT="T" AND '+
               'C_AUXILIAIRE="'+StAuxi +'"');
    //mcd 02/03/2005 manque adresse et tiersfrais
   ExecuteSQL ('Update ADRESSES set adr_natureauxi="CLI" where ADR_REFCODE="'+StAuxi+'"');
   ExecuteSQL ('Update TIERSFRAIS set gtf_natureauxi="CLI" where gtf_tiers="'+StAuxi+'"');
   if (GetParamsoc('SO_AFLIENDP') =true)  then  //mcd 06/10/2005 il faut changer dans annuaire
       ExecuteSQL ('Update ANNUAIRE set ann_natureauxi="CLI" where ann_tiers="'+CodeTiers+'"');
end;

{***********A.G.L.***********************************************
Auteur  ...... : G.Merieux
Créé le ...... : 10/01/2003
Modifié le ... : 10/01/2003
Description .. : Recherche adresse facturation si le tiers facturé est
Suite ........ : personnalisé dans l'affaire
Mots clefs ... : ADRESSE;GIGA
*****************************************************************}
Procedure RechAdresseFactAffaire ( TOBTiers,TOBAdresses,TOBPiece,TOBPieceAffaire : TOB ) ;
Var AuxiLivr,CodeLivr,AuxiFact,CodeFact,AffFact : String ;
BEGIN
if TOBTiers=Nil then Exit ;
CodeLivr:=TOBTiers.GetValue('T_TIERS') ; if CodeLivr='' then Exit ;
AuxiLivr:=TOBTiers.GetValue('T_AUXILIAIRE') ;
AuxiFact:=TOBTiers.GetValue('T_FACTURE') ;
if ((AuxiFact='') or (AuxiFact=AuxiLivr)) then BEGIN AuxiFact:=AuxiLivr ; CodeFact:=CodeLivr ; END
                                          else BEGIN CodeFact:=TiersAuxiliaire(AuxiFact,True) ; END ;
// specif Affaire , à voir pour remettre dans la fonction    TiersVersAdresse
	if (TOBPieceAffaire = NIL) then  exit;

  AffFact := TOBPieceAffaire.GetValue('GP_TIERSFACTURE');
  if (AffFact<>'') then CodeFact := AffFact;     // on prends le cli a facturer de l'affaire


// Facturation
if CodeFact=CodeLivr then
   BEGIN
     GetAdrFromTOB(TOBAdresses.Detail[1],TOBTiers) ;
     if GetParamSoc('SO_GCPIECEADRESSE') then TOBPiece.PutValue('GP_NUMADRESSEFACT',+1)
                                         else TOBPiece.PutValue('GP_NUMADRESSEFACT',-1);
   END
else
   BEGIN
   GetAdrFromCode(TOBAdresses.Detail[1],CodeFact) ;
   if GetParamSoc('SO_GCPIECEADRESSE') then   TOBPiece.PutValue('GP_NUMADRESSEFACT',+2)
   																			else   TOBPiece.PutValue('GP_NUMADRESSEFACT',-2);
   END ;
END ;

Procedure AFLignePieceToTache ( pStCodeAffaire,pStCodeTiers,pStLignePieceAff : String; pAction:TActionFiche );
var vStArg,vStCle : string;
begin
  vStCle := findEtreplace (pStLignePieceAff ,';','|',true);
  vStArg := 'ATA_AFFAIRE:'+ pStCodeAffaire +';ATA_TIERS:'+ pStCodeTiers +';ATA_IDENTLIGNE:'+ vStCle +';';
  vStArg := vStArg + ActionToString(pAction);
  AGLLanceFiche ('AFF','AFTACHES','', '',vStArg);
end;

function IsTypeAffaire (Affaire : string; TypeAffaire : string) : boolean;
begin
	result :=  (copy (Affaire,1,1)=TypeAffaire);
end;

// OPTIMISATIONS LS
procedure ReinitTOBAffaires;
begin
  TOBLesAffaires.ClearDetail;
end;

function FindCetteAffaire (CodeAffaire : string) : TOB;
begin
  result := TOBlesAffaires.findFirst(['AFF_AFFAIRE'],[CodeAffaire],true);
end;

procedure StockeCetteAffaire (CodeAffaire : string);
var QQ : TQuery;
    TOBLocal : TOB;
begin
  if TOBlesAffaires.findFirst(['AFF_AFFAIRE'],[CodeAffaire],true) = nil then
  begin
    QQ := OpenSql ('SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE="'+CodeAffaire+'"',true,-1,'',true);
    TOBlocal := TOB.Create ('AFFAIRE',nil,-1);
    TOBLocal.selectDb ('',QQ);
    ferme(QQ);
    TOBLocal.changeParent (TOBLesAffaires,-1);
  end;
end;

procedure StockeCetteAffaire (TOBAffaire : TOB);
var TOBLocal : TOB;
begin
  if TOBlesAffaires.findFirst(['AFF_AFFAIRE'],[TOBaffaire.getValue('AFF_AFFAIRE')],true) = nil then
  begin
    TOBLocal := TOB.create ('AFFAIRE',nil,-1);
    TOBlocal.Dupliquer (TOBAffaire,false,true);
    TOBLocal.changeParent (TOBLesAffaires,-1);
  end;
end;

function ControleDocumentsAffaire (CodeAffaire : string) : boolean;
var QQ,Q1 : TQuery;
		TOBPIECES,TOBP: TOB;
    Indice : integer;
    nature : string;
    okok : boolean;
begin
	result := false;
  okOk := true;
  TOBPIECES := TOB.create ('LES PIECES',nil,-1);
  QQ := OpenSql ('SELECT * FROM PIECE WHERE GP_AFFAIRE="'+CodeAffaire+'" AND GP_VENTEACHAT="VEN"',true,-1,'',true);
  TRY
    if not QQ.eof then
    begin
      TOBPIECES.LoadDetailDB ('PIECE','','',QQ,false,true);
      for Indice := 0 to TOBPIECES.detail.count -1 do
      begin
        TOBP := TOBPIECES.detail[Indice];
        Nature := TOBP.getValue('GP_NATUREPIECEG');
        if (nature <> 'DBT') And (Nature <> 'AFF') then
        begin
          PgiInfo ('Impossible des documents d''exploitation sont en cours');
          okok := false;
          break;
        end;
        if (nature = 'DBT') then
        begin
        	Q1 := OpenSql ('SELECT AFF_ETATAFFAIRE FROM AFFAIRE WHERE AFF_AFFAIRE="'+
          								TOBP.GetValue('GP_AFFAIREDEVIS')+'"',true,1,'',true);
          if not Q1.eof then
          begin
          	if (Q1.findField('AFF_ETATAFFAIRE').AsString = 'ACP') then
            begin
          		PgiInfo ('Impossible un devis de l''affaire est accepté');
            	okok := false;
              break;
            end;
          end;
          ferme (Q1);
        end;
      end;
    end;
    //if Okok then result := true;
  FINALLY
  	ferme (QQ);
    FreeAndNil(TOBPIECES);
  END;

  //contrôle si des appels sont affectés à ce contrat
  if okok then
  begin
    if (pos('I', codeAffaire) <> 0) then
    begin
    Q1 := OpenSQL ('SELECT AFF_AFFAIREINIT FROM AFFAIRE WHERE AFF_AFFAIREINIT="'+ CodeAffaire+'"',true,1,'',true);
    if Not Q1.Eof then
    begin
      PgiInfo ('Impossible des appels sont en cours sur ce contrat');
      okok := false;
    end;
    ferme (Q1);
    end;
  end;

  result := OkOk;

end;


function CalculMtEcheanceContrat (MtDoc: double ;CalcEcheance,stPeriode: string;Interval,NbIt : integer) : double;
begin
  if CalcEcheance = 'DME' then result := mtDoc //Montant Du Document = Montant Echéance
  Else if CalcEcheance = 'DMM' then //Montant du document = Montant au Mois
  Begin
     if StPeriode = 'NBI' then
        result := MtDoc
     Else if StPeriode = 'M' then
        result := MtDoc * Interval
     Else if Stperiode = 'S' then
        result := (Mtdoc / 4) *  Interval
     Else if Stperiode = 'A' then
        result := (Mtdoc * 12);
  end
  else if CalcEcheance = 'DMP' then  //Montant du Document = Montant Période de Fact.
       result := MtDoc / NbIt
  else if CalcEcheance = 'DMA' then //Montant Du Document = Montant Annuel
  Begin
     if StPeriode = 'M' then
        result := (Mtdoc / 12) *  Interval
     Else if StPeriode = 'S' then
        result := (Mtdoc / 52) *  Interval
     else if StPeriode = 'A' then
        Result := MtDoc;
  end;
end;



INITIALIZATION
    TOBlesAffaires := TOB.create ('LES AFFAIRES',nil,-1);
FINALIZATION
    TOBlesAffaires.free;
//
end.

