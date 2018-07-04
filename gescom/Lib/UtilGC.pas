unit UtilGC;

interface

Uses
{$IFDEF EAGLCLIENT}
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, db, {$ELSE} uDbxDataSet, {$ENDIF}Dialogs,
{$IFDEF V530}
     EdtEtat,
{$ELSE}
     EdtREtat,
{$ENDIF}
{$ENDIF}
     Ent1, Sysutils, HCtrls, HMsgBox, Forms, StdCtrls, HEnt1, EntGC, UtilArticle,Utob,
     ParamSoc,M3FP, Hstatus, UtilTOM,Controls,UtilPgi,uEntCommun,Hdb ;

// Modif BTP
Function FiltreNonascii (Zone : string) : string;
Function AttribNewCodeB( NomTable:string; NomChamp:string; LgMaxi:integer; Prefixe:string;  suffixe:string) : string ;
// ----
Function  AttribNewCode( NomTable:string; NomChamp:string; LgMaxi:integer; Prefixe:string; DernierChrono:string; TypeCAB:string; CtrlNonExiste : boolean=True; AlternativeBTP : Boolean=false) : String ;
Function  MajDernierChrono(NomChamp : string) : integer;
Function  ExtraitChronoCode( Code:string ) : String ;
Function  AttributCodeComptable( Code:string ) : String ;
function  RechParamCAB_Auto( FournPrinc : String; var v_TypeCAB : String; Force : boolean=False ) : boolean ;
Function  AttribNewCodeCAB( CtrlNonExiste : boolean ; FournPrinc : string = '') : string ;
procedure SauvChronoCAB( FournPrinc : String ) ;
function  RechParamCAB_Manu( FournPrinc : String; var v_TypeCAB:String; var v_LgNumCAB:integer ) : boolean ;
function  ExtractLibelle ( St : string) : string;
function  SpaceStr ( nb : integer) : string;
{$ifdef GCGC}
Function RTPhonemeSearch(St1 : String) : String ;
{$endif}
Procedure EditMonarch (Monarch : String);
Procedure EditMonarchSiEtat (Etat : String);
function EtatMonarchFactorise (Etat : String): boolean;
Function CompteNbEtablissMode : Integer ;
{$IFNDEF EAGL}
Procedure LanceEtatLibreGC (CodNature:String='');
{$ENDIF}

{$IFDEF GCGC}
Function  GCMAJChampLibre (FF:TForm; Edition:Boolean; TypeChp:String; Name:String; Nbre:Integer; Suffixe:String=''; AvecPrefixe:boolean=false; LePrefixe:string='') : integer;
Function  GCXXWhereChampLibre (FF:TForm; xx_where:String; TypeChp:String; Name:String; Nbre:Integer; Suffixe:String='') : string;
Function  GCTitreZoneLibre (Name : String; Var TitreZone : String; AvecPrefixe : boolean=false;  LePrefixe:string=''): Boolean;
Function  CodeTableLibre_2 (Name : String; CLiFou : string = ''): string;
Function  LibTableLibreCliFour(VenteAchat, NomDuChamp : string) : string;
Function  GRCCodeTableLibre (Name : String): string;
Procedure GCModifTitreZonesLibres (Prefixe : String; stCode : string = 'ZLI') ;
Procedure GCModifLibelles;
Procedure GCTripoteStatus ;
Function  ChangeLibelleCombo (t,s : string) : string ;
Function ChangeBoolLibre ( NomChamp : string ; FF : Tform; SansUnderScore : boolean=false ): boolean;
Function ChangeBoolLibre2 ( NomChamp,Nomzone : string ; FF : Tform; SansUnderScore : boolean=false ): boolean;
//Procedure MajChampsLibresArticle(FF: Tform; TSTablesLibres:string='PTABLESLIBRES';TSZonesLibres : string='PZONESLIBRES' );
Procedure MajChampsLibresArticle(FF: Tform; stPrefixe:string='GA';MasqueSiVide:boolean=true;TSTablesLibres:string='PTABLESLIBRES';TSZonesLibres : string='PZONESLIBRES' );
  //mcd 07/07/03
function AfPlusNatureAchat(BLC:boolean=FALSE;ForPlus : boolean=true) :string;
{$ENDIF}
procedure GCTestDepotOblig;
  Procedure MajChampsLibresGED( FF : Tform ) ;
function MAJJnalEvent(TypeEvt, Etat, Libelle, BlocNote : string; FileName: string ='') : integer;
function TransformeLesInToOr(stIn : string) : string;
function GereCommercial : boolean;
procedure TraiteParametresSurTForm(FF : TForm; LesChamps : string=''; AffecteFiltre : boolean=false);
function NbOccurenceString(Chaine, Occurence : string) : integer;
Procedure RTCreerProspectPourClient ;
Procedure RTAlimCleTelephonTiers ;
Procedure RTAlimTelephonContact ;
function IntervertirNomChampsXX ( FF:Tform; sChampSansIndice:string; iDebut, iFin:integer; RendInvisible:boolean=false) : integer;
Function ChangeLibre2 ( NomChamp : string ; FF : Tform ) : boolean;
Function RTTypeProduitCRM : string;

Type T_SupArt = Class
                   // DCA - FQ MODE 10815 - Ajout du statut
                   Cle, Code, Statut, TypeArt : String ;
                   Procedure SuppressionArticle ;
                   End ;

implementation

uses
  {$IFDEF GPAO}
    wConversionUnite,
  {$ENDIF}
{$IFDEF BTP}
     UtilsMetresXLS,
{$ENDIF}
  Choix,
  Printers
  ,CbpMCD
  ,CbpEnumerator
  ,UFileAssoc
  ;

const
	// libellés des messages
	TexteMessage: array[1..3] of string 	= (
          {1}        'Attribution du code'
          {2}        ,'Le Préfixe du code a été racourci'
          {3}        ,'La longueur maxi du code a été atteinte'
                         );

//
// Déclaration de variables locales
//

Var PrefixeCAB : String ;     // Préfixe du code à barres
    TypeCAB    : String ;     // Type du code à barres
    ChronoCAB  : String ;     // Chrono correspondant au dernier code à barres généré
    ChronoCode : String ;     // Chrono correspondant au dernier code généré (Tiers, Article ou CAB)
    LgNumCAB   : integer;     // Longueur du code à barres
    CABAuto    : boolean;     // Attribution automatique d'un code à barres
    CABAutoFour: boolean;     // Attribution automatique d'un code à barres du fournisseur

{***********A.G.L.***********************************************
Auteur  ...... : LS
Créé le ...... : 10/08/2001
Modifié le ... :
Description .. : Attribution automatique d'un code Article
Suite ........ : à partir de son libellé et d'un préfixe
Mots clefs ... : LIBELLE;CODE ARTICLE;
*****************************************************************}
Function AttribNewCodeB( NomTable:string; NomChamp:string; LgMaxi:integer; Prefixe:string; suffixe:string) : string ;
var IndChrono,LgChrono, Lgsuffixe : Integer ;
    CodeOk : Boolean;
    NewCode,Ssuffixe : string;
BEGIN
NewCode:='';
IndChrono := 0;
LgChrono := 0;
CodeOk := False;
if length(Prefixe) >= LgMaxi then begin Prefixe:=''; end;
suffixe := FiltreNonascii (suffixe);
While not CodeOk do
begin
  lgsuffixe := LgMaxi - length (Prefixe) - LgChrono;
  Ssuffixe := copy (suffixe,1,lgsuffixe);
  if LgChrono > 0 then
     Ssuffixe := Ssuffixe + Inttostr (IndChrono);
  // Recherche si existe déjà
  NewCode:= Prefixe + Ssuffixe;
  if (not ExisteSQL('Select '+NomChamp+' From '+NomTable+' Where '+NomChamp+'="'+NewCode+'"')) then
     CodeOk := True
  else
  begin
     IndChrono := IndChrono + 1;
     LgChrono := length (inttostr(Indchrono));
  end;
end;
Result:= NewCode;
END;

{***********A.G.L.***********************************************
Auteur  ...... : JCF
Créé le ...... : 15/10/2001
Modifié le ... : 27/05/2002 par TC - utilisation de plusieurs Monarch
Description .. : Gestion des formats différents des imprimantes MONARCH
Suite ........ : <Nom Monarch> pour éditer sur une Monarch
Suite ........ : '' pour éditer sur l'imprimante par défaut
Mots clefs ... : EDITION;ETIQUETTE;MONARCH;ZEBRA;
*****************************************************************}
Procedure EditMonarch (Monarch : string);
var cpt : integer;
    FPrinter: TPrinter;
BEGIN
if Monarch <>''then
   begin
   FPrinter := Printer;
   For cpt := 0 to FPrinter.printers.count - 1 do
       begin
       if pos (Uppercase(Monarch), Uppercase (Printer.printers[cpt])) > 0 then
          begin
          V_PGI.QRPrinterindex := cpt;
          Break;
          end;
       end;
   end
   else begin
        V_PGI.QRPrinterindex := -1;            // Imprimante par défaut
        end;
END;

Procedure EditMonarchSiEtat (Etat : String);
var debutChaineEtat,finChaineEtat,nomMonarch : string;
BEGIN
if pos ('MONARCH', Uppercase (Etat)) > 0 then
   begin
       // je récupère le nom de l'état à partir du mot 'MONARCH'
   debutChaineEtat:=copy(Etat,pos('MONARCH',Uppercase(Etat)),length(Etat));
      // je récupère La chaine qui suis le nom complet de la monarch
   finChaineEtat := StrScan(PChar(debutChaineEtat),' ');
     // je récupère le nom complet de la monarch
   nomMonarch:=copy(debutChaineEtat,1,Length(debutChaineEtat)-Length(finChaineEtat));
   EditMonarch (nomMonarch);
   end else
if pos ('ZEBRA', Uppercase (Etat)) > 0 then
   begin
       // je récupère le nom de l'état à partir du mot 'ZEBRA'
   debutChaineEtat:=copy(Etat,pos('ZEBRA',Uppercase(Etat)),length(Etat));
      // je récupère La chaine qui suis le nom complet de la zebra
   finChaineEtat := StrScan(PChar(debutChaineEtat),' ');
     // je récupère le nom complet de la zebra
   nomMonarch:=copy(debutChaineEtat,1,Length(debutChaineEtat)-Length(finChaineEtat));
   EditMonarch (nomMonarch);
   end else
   EditMonarch('');
END;

{***********A.G.L.***********************************************
Auteur  ...... : TC
Créé le ...... : 09/07/2002
Modifié le ... :
Description .. : Permet de savoir si c'est état pour imprimante monarch
Suite ........ : qui imprime de façon factorisé
Mots clefs ... : MONARCH,FACTORISE
*****************************************************************}

function EtatMonarchFactorise (Etat : String) : boolean;
BEGIN
result:=False;
if (pos ('MONARCH', Uppercase (Etat)) > 0) or  (pos ('ZEBRA', Uppercase (Etat)) > 0) then
   begin
   if copy(Etat,length(Etat)-1,length(Etat))=' F' then  result:=True;
   end ;
END;

{***********A.G.L.***********************************************
Auteur  ...... : JCF
Créé le ...... : 20/12/2001
Modifié le ... : 20/12/2001
Description .. : Rend le nombre d'établissements de la base
Suite ........ : au sens MODE (nombre géré sur site)
Mots clefs ... : ETABLISS;SERIA;MODE
*****************************************************************}
Function CompteNbEtablissMode : Integer ;
var QQ :TQuery ;
begin
Result:=0 ;
if (ctxMode in V_PGI.PGIContexte) then
	begin
    QQ:=OpenSQL('SELECT COUNT(*) FROM ETABLISS WHERE ET_SURSITE="X"',True) ;
    if Not QQ.EOF then Result:=QQ.Fields[0].AsInteger ;
    Ferme(QQ) ;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : JT
Créé le ...... : 04/07/2003
Modifié le ... : 04/07/2003
Description .. : Calcul "sécurisé" du chrono tiers ou article
Mots clefs ... : CHRONO;CODE BARRE;
*****************************************************************}
Function MajDernierChrono(NomChamp : string) : integer;
var Cpt, Maj, iChrono : integer;
    Qry : TQuery;
    Ok : boolean;
    s, Nomtable,OldValeur : String;
begin
// Modif BRL 16/02/2010
// récupération base principale en cas de partage du compteur
  NomTable := 'PARAMSOC';
  Qry := OpenSQL('SELECT DS_NOMBASE FROM DESHARE WHERE DS_NOMTABLE = "'+NomChamp+'"',False);
  if Not Qry.Eof then
  Begin
    NomTable := Qry.FindField('DS_NOMBASE').AsString+'.DBO.PARAMSOC';
  End;
  Ferme(Qry);

  Result := 0;
  Ok := False;
  Cpt := 0;
  repeat
    inc(Cpt);
    Qry := OpenSQL('SELECT SOC_DATA FROM '+Nomtable+' WHERE SOC_NOM = "'+NomChamp+'"',False);
    OldValeur := Qry.FindField('SOC_DATA').AsString;
    iChrono := Qry.FindField('SOC_DATA').AsInteger;
    Ferme(Qry);
    Maj := ExecuteSQL('UPDATE '+Nomtable+' SET SOC_DATA = "'+IntToStr(iChrono + 1)+'" WHERE SOC_NOM = "'+NomChamp+'" '+
                      'AND SOC_DATA ="'+Oldvaleur+'"');
    Ok := (Maj = 1);
  until ((Ok) or (Cpt > 20));

  if Ok then
  begin
    ChronoCode := IntToStr(iChrono);
    Result := iChrono+1;
  end else
  begin
    ChronoCode := '';
    Result := 0;
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : AV
Créé le ...... : 22/12/2000
Modifié le ... : 16/08/2001
Description .. : Attribution automatique d'un code Tiers, d'un code Article
Suite ........ : ou
Suite ........ : d'un Code à barres, à partir du premier chrono disponible.
Suite ........ : CtrlNonExiste permet de contrôler ou non la non existence
Suite ........ : du code attribué
Mots clefs ... : CHRONO;CODE BARRE;
*****************************************************************}
Function AttribNewCode( NomTable:string; NomChamp:string; LgMaxi:integer; Prefixe:string; DernierChrono:string; TypeCAB:string; CtrlNonExiste : boolean=True; AlternativeBTP : Boolean=false) : string ;
var i, LgPrefixe, LgChrono : Integer ;
    i_Chrono,i_Reste: Int64;
    NbBoucle : Integer;
    CodePrefixe, NewCode : string;
    Modulo10 : Boolean ;
BEGIN
  NbBoucle:=0;
  NewCode:='';
  ChronoCode:=DernierChrono;

  if (NomTable='TIERS') and (LgMaxi=-1) then Modulo10:=True else Modulo10:=false;

  if  (TypeCAB='E13') or
      (TypeCAB='EA8') or
      (TypeCAB='39C') or
      (TypeCAB='ITC') or
      (TypeCAB='ITF') then
      LgMaxi:=LgMaxi-1;                  //Clé de contrôle sur le dernier caractère

  if (LgMaxi < 4) and (LgMaxi>0) then LgMaxi:=VH^.Cpta[fbAux].Lg;

  LgPrefixe := Length(Prefixe);
  if (LgPrefixe >= LgMaxi) and (LgMaxi>0) then
  begin
  {LgPrefixe:=0;}
  Prefixe:='';
  end;

  While NbBoucle <= 3 do
  begin
  {Si chrono TIERS, appel génération "sécurisée"}
  if (NomTable='TIERS') then
  begin
    if AlternativeBTP then
    begin
    	i_Chrono := MajDernierChrono('SO_BTCOMPTEURTIERS');
      end
      else
    begin
    	i_Chrono := MajDernierChrono('SO_GCCOMPTEURTIERS');
    end;
      //
    if i_Chrono = 0 then
    begin
      Result:= '';
      exit;
    end;
    end
    else
  begin
    // DCA - RHA 16/10/2003 Création du code barre nomenclature
    if ChronoCode = '' then ChronoCode := '0';
    i_Chrono := StrToInt64(ChronoCode);
    i_Chrono := i_Chrono+1;
  end;
    //
  if (Modulo10) then
     begin
     i_Reste:=i_Chrono Mod 10 ;
     if i_Reste = 0 then i_Chrono := i_Chrono+1;
     end;
  ChronoCode := IntToStr(i_Chrono);
  LgChrono := Length(ChronoCode);
  CodePrefixe := Prefixe;
  LgPrefixe := Length(CodePrefixe);
  if ((LgPrefixe+LgChrono) > LgMaxi) and (NbBoucle < 3) and (LgMaxi>0) then // on est au bout du compteur
     begin
     if NbBoucle = 0 then                         // remise du compteur à 0
        begin
        ChronoCode:='1';
        LgChrono:=1;
        NbBoucle:=1;
      end
      else if (NbBoucle = 1) and (LgPrefixe > 0) then   // Si préfixe, on le racourci
        begin
        LgPrefixe:=LgPrefixe-1;
        CodePrefixe:=Copy(Prefixe, 1, LgPrefixe);
        Prefixe:=CodePrefixe;
        PgiBox (TexteMessage[2], TexteMessage[1]);
        NbBoucle:=2;
      end
      else
        begin                                     // Suppression du préfixe
        LgPrefixe:=0;
        CodePrefixe:='';
        PgiBox (TexteMessage[3], TexteMessage[1]);
        NbBoucle:=3;
        end;
     end;
    //                                                 // Recherche si existe déjà
   for i:=LgPrefixe to (LgMaxi - LgChrono - 1) do CodePrefixe:=CodePrefixe+'0' ;
   //LgPrefixe := Length(CodePrefixe);
   NewCode:= CodePrefixe + ChronoCode;
   if (TypeCAB='E13') or (TypeCAB='EA8') or (TypeCAB='39C') or (TypeCAB='ITC') or (TypeCAB='ITF') then
      begin
      // Ajout de la clé de contrôle
      NewCode := NewCode + CalculCleCodeBarre(NewCode+'0', TypeCAB );
      end;
   if (not CtrlNonExiste) then NbBoucle:=8
   else
     begin
     if NomChamp='GA_CODEARTICLE' then //Champ indexé dans la table article
       begin
       if (not ExisteSQL('Select GA_CODEARTICLE From ARTICLE Where GA_ARTICLE="'+CodeArticleUnique2(NewCode,'')+'"')) then NbBoucle:=9;
       end
     else
       begin
       if (not ExisteSQL('Select '+NomChamp+' From '+NomTable+' Where '+NomChamp+'="'+NewCode+'"')) then NbBoucle:=9;
       end;
     end;
  end;
Result:= NewCode;
END;

{***********A.G.L.***********************************************
Auteur  ...... : AV
Créé le ...... : 22/12/2000
Modifié le ... :
Description .. : Extraction du chrono à partir d'un code Tiers ou d'un code Article
Suite ........ :
Mots clefs ... : CHRONO;
*****************************************************************}
Function ExtraitChronoCode( Code:string ) : String ;
var LgPrefixe, LgCode, i : Integer ;
    i_Chrono : Int64;
    Chrono : string;
BEGIN
Code := Trim(Code);
LgCode := Length(Code);
LgPrefixe := LgCode;
For i:=LgCode downto 1 do
    if (Code[i]<'0') or (Code[i]>'9')
       then break
       else LgPrefixe:=i-1;
Chrono := Copy(Code, LgPrefixe+1, LgCode-LgPrefixe);
i_Chrono := StrToInt64(Chrono);
Result := IntToStr(i_Chrono);
END;

{***********A.G.L.***********************************************
Auteur  ...... : AV
Créé le ...... : 22/12/2000
Modifié le ... :
Description .. : Attribution du code comptable Tiers par défaut à partir du code Tiers.
Suite ........ : La fonction supprime ou rajoute des zéro à gauche du chrono Tiers, si
Suite ........ : la longueur du code comptable est différente de celle du code tiers.
Mots clefs ... : CHRONO
*****************************************************************}
Function AttributCodeComptable( Code:string ) : String ;
var LgPrefixe, LgCode, LgCpte, LgChrono : Integer ;
    i : Integer;
    NumChrono : Int64;
    Chrono, CodePrefixe : string;
BEGIN
Result := Code;
LgCode := Length(Trim(Code));
LgCpte := VH^.Cpta[fbAux].Lg;
if (lgCode <> LgCpte) then
   begin
   LgPrefixe := LgCode;
   For i:=LgCode downto 1 do
    if (Code[i]<'0') or (Code[i]>'9')
       then break
       else LgPrefixe:=i-1;
   CodePrefixe := Copy(Code, 1, LgPrefixe);
   Chrono := Copy(Code, LgPrefixe+1, LgCode-LgPrefixe);
   NumChrono := StrToInt64(Chrono);
   Chrono := IntToStr (NumChrono);
   LgChrono := Length(Chrono);
   for i:=LgPrefixe to (LgCpte - LgChrono - 1) do CodePrefixe:=CodePrefixe+'0' ;
   Result:= CodePrefixe + Chrono;
   end;
END;

//
// Fonction de recherche du paramétrage des codes à barres (Attribution automatique)
// Le paramètre Force à True permet de récupérer le paramétrage du CAB même si le paramètre
// société SO_GCNUMCABAUTO vaut False (utile dans le module de génération des CAB)
//
function RechParamCAB_Auto( FournPrinc : String; var v_TypeCAB : String; Force : boolean=False ) : boolean ;
Var QQ : TQuery ;
begin
PrefixeCAB:='';
TypeCAB:='';
ChronoCAB:='0' ;
LgNumCAB:=0;
CABAuto:=False;
CABAutoFour:=False;
if (GetParamsoc('SO_GCNUMCABAUTO') or Force) then
  begin
  CABAuto:=True;
  if (GetParamsoc('SO_GCCABFOURNIS')=True) and (FournPrinc <> '') then
     begin
     // Recherche du paramétrage du CAB par rapport au fournisseur de l'article
     QQ:=OpenSQL('Select * from CODEBARRES where (GCB_NATURECAB="T") and (GCB_IDENTIFCAB="'+FournPrinc+'")',TRUE);
     If Not QQ.EOF then
        begin
        if QQ.Findfield('GCB_CABTIERS').AsString='X' then
           begin
           if QQ.Findfield('GCB_NUMCABAUTO').AsString='X' then
              begin
              CABAutoFour:=True;    // Code à barres auto. spécifique au fournisseur
              LgNumCAB:=QQ.Findfield('GCB_LGNUMCAB').AsInteger;
              PrefixeCAB:=QQ.Findfield('GCB_PREFIXECAB').AsString;
              ChronoCAB:=QQ.Findfield('GCB_COMPTEURCAB').AsString;
              TypeCAB:=QQ.Findfield('GCB_TYPECAB').AsString;
              end
              else CABAuto:=False;  // Code à barres manuel
           end;
        end;
     ferme(QQ) ;
     end;
  if (CABAuto=True) and (CABAutoFour=False) then
     begin
     // Recherche du paramétrage global du Code à barres
     QQ:=OpenSQL('Select * from CODEBARRES where (GCB_NATURECAB="SO") and (GCB_IDENTIFCAB="...")',TRUE);
     If Not QQ.EOF then
        begin
        LgNumCAB:=QQ.Findfield('GCB_LGNUMCAB').AsInteger;
        PrefixeCAB:=QQ.Findfield('GCB_PREFIXECAB').AsString;
        ChronoCAB:=QQ.Findfield('GCB_COMPTEURCAB').AsString;
        TypeCAB:=QQ.Findfield('GCB_TYPECAB').AsString;
        end
        else CABAuto:=False;
     ferme(QQ) ;
     end;
  end;
  Result := CABAuto ;
  v_TypeCAB := TypeCAB ;
end;

//
// Fonction d'attribution d'un nouveau code à barres
//
Function AttribNewCodeCAB( CtrlNonExiste : boolean ; FournPrinc : string = '') : string ;
var ChronoCABInit : string ;
    bOkChrono : boolean ;
    iOcc : integer ;
begin
  iOcc := 0 ;
  // Gestion du multi-utilisateurs : maj CODEBARRES avec valeur de l'ancien chrono
  repeat
    inc ( iOcc ) ;
    ChronoCABInit := ChronoCAB ;
    Result:=AttribNewCode('ARTICLE','GA_CODEBARRE',LgNumCAB,PrefixeCAB,ChronoCAB,TypeCAB,CtrlNonExiste) ;

    if not CABAutoFour then
         bOkChrono := ExecuteSQL('update CODEBARRES set GCB_COMPTEURCAB="'+ChronoCode+
           '" where (GCB_NATURECAB="SO") and (GCB_IDENTIFCAB="...") and GCB_COMPTEURCAB="'+ChronoCABInit+'"') = 1
    else bOkChrono := ExecuteSQL('update CODEBARRES set GCB_COMPTEURCAB="'+ChronoCode+
           '" where (GCB_NATURECAB="T") and (GCB_IDENTIFCAB="'+FournPrinc+'") and GCB_COMPTEURCAB="'+ChronoCABInit+'"' ) = 1 ;

    // Rechargement du dernier chrono utilisé : récup. ChronoCAB
    if not bOkChrono then RechParamCAB_Auto ( FournPrinc, TypeCAB) ;

  until (( bOkChrono ) or ( iOcc > 20 )) ;

  if not bOkChrono then Result := '' ;

  // Récupération du nouveau chrono
  ChronoCAB := ChronoCode ;
end;

//
// Sauvegarde du dernier chrono affecté aux codes à barres
//
procedure SauvChronoCAB( FournPrinc : String ) ;
begin
  if CABAutoFour=False
     then ExecuteSQL('UPDATE CODEBARRES SET GCB_COMPTEURCAB="'+ChronoCAB+
                     '" WHERE (GCB_NATURECAB="SO") and (GCB_IDENTIFCAB="...")')
     else ExecuteSQL('UPDATE CODEBARRES SET GCB_COMPTEURCAB="'+ChronoCAB+
                     '" WHERE (GCB_NATURECAB="T") and (GCB_IDENTIFCAB="'+FournPrinc+'")') ;
end;


//
// Fonction de recherche du paramétrage des codes à barres (Attribution Manuelle)
//
function RechParamCAB_Manu( FournPrinc : String; var v_TypeCAB:String; var v_LgNumCAB:integer ) : boolean ;
Var QQ : TQuery ;
    CABManu : Boolean ;
begin
CABManu:=False;

  if (GetParamsoc('SO_GCCABFOURNIS')=True) and (FournPrinc <> '') then
     begin
     // Recherche du paramétrage du CAB par rapport au fournisseur de l'article
     QQ:=OpenSQL('Select * from CODEBARRES where (GCB_NATURECAB="T") and (GCB_IDENTIFCAB="'+FournPrinc+'")',TRUE);
     If Not QQ.EOF then
        begin
        if QQ.Findfield('GCB_CABTIERS').AsString='X' then
           begin
           if QQ.Findfield('GCB_NUMCABAUTO').AsString='-' then
              begin
              CABManu:=True;    // Code à barres manu. spécifique au fournisseur
              v_LgNumCAB:=QQ.Findfield('GCB_LGNUMCAB').AsInteger;
              v_TypeCAB:=QQ.Findfield('GCB_TYPECAB').AsString;
              end;
           end;
        end;
     ferme(QQ) ;
     end;
  if (CABManu=False) and (GetParamsoc('SO_GCNUMCABAUTO')=False) then
     begin
     // Recherche du paramétrage global du Code à barres
     QQ:=OpenSQL('Select * from CODEBARRES where (GCB_NATURECAB="SO") and (GCB_IDENTIFCAB="...")',TRUE);
     If Not QQ.EOF then
        begin
        CABManu:=True;
        v_LgNumCAB:=QQ.Findfield('GCB_LGNUMCAB').AsInteger;
        v_TypeCAB:=QQ.Findfield('GCB_TYPECAB').AsString;
        end;
     ferme(QQ) ;
     end;
  Result := CABManu ;
end;

//////////////////// Suppression d'un article ////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Michel RICHAUD
Créé le ...... : 19/07/2001
Modifié le ... : 19/07/2001
Description .. : Suppression d'un article implémentée par DB
Mots clefs ... : SUPPRESSION;ARTICLE
*****************************************************************}
Procedure T_SupArt.SuppressionArticle;
var stWhereArt : string ;
    stInNomen : string;
{$IFDEF BTP}
	 MetreArticle : TMetreArt;
{$ENDIF}
begin

  // DCA - FQ MODE 10815 - Gestion suppression article dimensionné seul
  if Statut = 'GEN'
  // Suppression de tous les articles dimensionnés et du générique
  // A priori IN() mieux que LIKE (???)
  then stWhereArt := ' IN (SELECT GA_ARTICLE FROM ARTICLE WHERE GA_STATUTART IN ("DIM","GEN") AND GA_CODEARTICLE="' + Code + '")'
  // Suppression de l'article unique ou du seul article dimensionné sélectionné
  else stWhereArt := '="' + Cle + '"' ;

ExecuteSQL('DELETE FROM TARIF WHERE GF_ARTICLE' + stWhereArt) ;
if GetPresentation=ART_ORLI then ExecuteSQL('DELETE FROM ARTICLECOMPL WHERE GA2_ARTICLE' + stWhereArt) ;
ExecuteSQL('DELETE FROM CONDITIONNEMENT WHERE GCO_ARTICLE' + stWhereArt) ;
ExecuteSQL('DELETE FROM ARTICLELIE WHERE GAL_ARTICLE' + stWhereArt) ;
ExecuteSQL('DELETE FROM DISPO WHERE GQ_ARTICLE' + stWhereArt) ;
ExecuteSQL('DELETE FROM ARTICLETIERS WHERE GAT_ARTICLE' + stWhereArt) ;
ExecuteSQL('DELETE FROM ACTIVITEPAIE WHERE ACP_ARTICLE' + stWhereArt) ;
ExecuteSQL('DELETE FROM RESSOURCEPR WHERE ARP_ARTICLE' + stWhereArt) ;
ExecuteSQL('DELETE FROM COMMISSION WHERE GCM_ARTICLE' + stWhereArt) ;
ExecuteSQL('UPDATE CATALOGU SET GCA_ARTICLE="" WHERE GCA_ARTICLE' + stWhereArt) ;
ExecuteSQL('DELETE FROM LIENSOLE WHERE LO_IDENTIFIANT' + stWhereArt) ;
ExecuteSQL('DELETE FROM GTRADARTICLE WHERE GTA_ARTICLE' + stWhereArt) ;
ExecuteSQL('DELETE FROM ARTICLEPIECE WHERE GAP_ARTICLE' + stWhereArt) ;
{$IFDEF BTP}
ExecuteSQL('DELETE FROM ARTICLECOMPL WHERE GA2_ARTICLE' + stWhereArt) ;
{$ENDIF}

// Correction BTP
stInNomen := ' IN (SELECT GNE_NOMENCLATURE FROM NOMENENT WHERE GNE_ARTICLE="' + cle + '")';
ExecuteSQL('DELETE FROM NOMENLIG WHERE GNL_NOMENCLATURE'+stInNomen); //MODIFBTP
//
ExecuteSQL('DELETE FROM NOMENENT WHERE GNE_ARTICLE="'+Cle+'"');

ExecuteSQL('DELETE FROM ARTICLE WHERE GA_ARTICLE' + stWhereArt) ;

{$IFDEF GPAO}
ExecuteSQL('DELETE FROM WCUCARA WHERE WDK_CONTEXTE="' + wContexteToString(wcuGA) + '" AND WDK_CUREFORIGINE="' + Cle + '"');
ExecuteSQL('DELETE FROM WCURESU WHERE WDR_CONTEXTE="' + wContexteToString(wcuGA) + '" AND WDR_CUREFORIGINE="' + Cle + '"');
{$ENDIF}

{$IFDEF BTP}
//Suppression des métres et variables de l'article
MetreArticle := TMetreArt.CreateArt(TypeArt, Cle);
MetreArticle.SupprimeVariable(Cle);
MetreArticle.SupprimeFichierXLS(MetreArticle.fRepMetre + '\Bibliotheque\' +  MetreArticle.FormatArticle(Cle) + '.Xlsx');
FreeAndNil(MetreArticle);
{$ENDIF}

end;

{***********A.G.L.***********************************************
Auteur  ...... : Michel RICHAUD
Créé le ...... : 19/07/2001
Modifié le ... : 19/07/2001
Description .. : Permet de récupérer le libellé d'un Label (ou autre) en
Suite ........ : passant la chaine correspondant au caption (ou autre) pour
Suite ........ : supprimer l'éventuel & et y ajoute deux points en fin.
Suite ........ : implémentée par DB
Mots clefs ... : LIBELLE;LABEL
*****************************************************************}
function ExtractLibelle ( St : string) : string;
Var St_Chaine : string ;
    i_pos : integer ;
begin
Result := '';
i_pos := Pos ('&', St);
if i_pos > 0 then
    begin
    St_Chaine := Copy (St, 1, i_pos - 1) + Copy (St, i_pos + 1, Length(St));
    end else St_Chaine := St;
Result := St_Chaine + ' : ';
end;

{***********A.G.L.***********************************************
Auteur  ...... : Michel RICHAUD
Créé le ...... : 19/07/2001
Modifié le ... : 19/07/2001
Description .. : Créé une chaine de blanc du nombre de caractères passé
Suite ........ : en
Suite ........ : paramètres
Mots clefs ... : CHAINE;BLANC
*****************************************************************}
function SpaceStr ( nb : integer) : string;
Var St_Chaine : string ;
    i_ind : integer ;
BEGIN
St_Chaine := '' ;
for i_ind := 1 to nb do St_Chaine:=St_chaine+' ';
Result:=St_Chaine;
END;

{$IFNDEF EAGL}
Procedure LanceEtatLibreGC (CodNature:String='');
Var CodeD, Nature : String ;
BEGIN
if CodNature <> '' then Nature:=CodNature
else if ctxFO in V_PGI.PGIContexte then Nature:='UFO'
else if ctxMode in V_PGI.PGIContexte then Nature:='UBO'
else exit ;
CodeD:=Choisir(TraduireMemoire('Choix d''un état libre'),'MODELES','MO_CODE || " " || MO_LIBELLE','MO_CODE','MO_NATURE="'+Nature+'" AND MO_MENU="X"','') ;
if CodeD<>'' then LanceEtat('E',Nature,CodeD,True,False,False,Nil,'','',False) ;
END ;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : LS
Créé le ...... : 05/04/2001
Modifié le ... : 05/04/2001
Description .. : Filtrage des codes non ascii
Suite ........ :
Mots clefs ... : FILTRE;CHAINE;ASCII
*****************************************************************}
Function FiltreNonascii (Zone : string) : string;
var
   Interm : string;
   Index : Integer;
begin
For Index := 1 to length (Zone) do
    begin
    case Zone[Index] of
    'a'..'z','A'..'Z','0'..'9':
         begin
         Interm := Interm + Zone[Index];
         end;
    'é','è','ê','ë' :
         begin
         Interm := Interm + 'e';
         end;
    'ä','à','â':
         begin
         Interm := Interm + 'a';
         end;
    'ù','û','ü','µ':
         begin
         Interm := Interm + 'u';
         end;
    'ö','ô':
         begin
         Interm := INterm + 'o';
         end;
    'ç':
         begin
         Interm:= Interm + 'c';
         end;
     end;
end;
Interm:=uppercase(Interm);
result := Interm;
end;

{$IFDEF GCGC}
{***********A.G.L.***********************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 07/07/2003
Modifié le ... :
Description .. : Fabrication de la condition plus des pièce d'achat
Suite ........ : Gérée en GI/GA
Mots clefs ... : NATUREPIECE;ACHAT;
*****************************************************************}
function AfPlusNatureAchat(BLC:boolean=FALSE;Forplus : boolean=true) :string;
begin

  // en GI on n'affiche toutes les natures non masquées (les natures non gérées sont masquées par l'utilisatire de création
  if ctxScot in V_PGI.PGIContexte  then result:=''
  // en GA, on renseigner les 4 natures gérées à ce jour (faire idem GI quand assistant création base existera ..)
 else
 begin
   if Forplus then
   begin
     if BLC then Result := ' AND GPP_NATUREPIECEG IN ("CF","BLF","FF","AF","BLC")'
            //FV1 : 08/01/2018 - FS#2832 - JPG - manque commandes regroupées dans consultations lignes
            else Result := ' AND GPP_NATUREPIECEG IN ("CF","CFR","LFR","BLF","FF","AF","AFS")';
            //else Result := ' AND GPP_NATUREPIECEG IN ("CF","BLF","FF","AF","AFS")';
   end
   else
   begin
     if BLC then Result := ' AND GP_NATUREPIECEG IN ("CF","BLF","FF","AF","BLC")'
            else Result := ' AND GP_NATUREPIECEG IN ("CF","BLF","FF","AF","AFS")';
 end;
 end;
end;



Function GCMAJChampLibre (FF:TForm; Edition:Boolean; TypeChp:String; Name:String; Nbre:Integer; Suffixe:String=''; AvecPrefixe:boolean=false; LePrefixe:string='') : integer;
Var Titre, Num, Suf : String;
{$IFDEF EAGLCLIENT}
    TChb  : TCheckBox;
    TCob  : THValComboBox;
    TLab  : THLabel;
    TEdi  : THEdit;
{$ELSE}
    TChb  : THDBCheckBox;
    TCob  : THDBValComboBox;
    TLab  : THLabel;
    TEdi  : THDBEdit;
{$ENDIF}
    i, j : Integer;
    Depart,Fin : integer;
begin
TLab:=Nil;
Result := Nbre;
if Pos('RSC_',Name) > 0 then
begin
	Depart := 0;
  Fin := Nbre-1;
end else
begin
	Depart := 1;
  Fin := Nbre;
end;
For i:=Depart to fin do
    begin
    if Depart = 1 then
    begin
    	if (i < 10) then Num := IntToStr(i) else Num := 'A';
    end else Num := IntToStr(i);
    GCTitreZoneLibre(Name+Num,Titre,AvecPrefixe,lePrefixe);
    if Titre = '' then Result := Result -1;
    if (TypeChp='BOOL') then
       begin
{$IFDEF EAGLCLIENT}
       TChb:=TCheckBox (FF.FindComponent (Name+Num)) ;
{$ELSE}
       TChb:=THDBCheckBox (FF.FindComponent (Name+Num)) ;
{$ENDIF}
       if TChb <> nil then TChb.Caption := '' ;
       if edition then TLab:=THLabel (FF.FindComponent ('T'+Name+Num)) ;
       if Titre = '' then
          begin
          if TChb <> nil then TChb.Visible := False ;
          if edition and (TChb <> nil) then TLab.Caption := '' ;
          end
       else if edition and (TLab <> nil) then TLab.Caption := Titre
       else if TChb <> nil then TChb.Caption := Titre ;
       {$IFDEF CCS3}
       if TChb<>Nil then TChb.Visible:=False ;
      {$ENDIF}
       end
    else if (TypeChp='COMBO') then for j:=0 to Length(Suffixe) do
        begin
        if j=0 then Suf:='' else Suf:=Suffixe[j];
{$IFDEF EAGLCLIENT}
        TCob:=THValComboBox (FF.FindComponent (Name+Num+Suf)) ;
{$ELSE}
        TCob:=THDBValComboBox (FF.FindComponent (Name+Num+Suf)) ;
{$ENDIF}
        TLab:=THLabel (FF.FindComponent ('T'+Name+Num+Suf)) ;
        if Titre = '' then
            begin
            if TCob <> nil then TCob.Visible := False ;
            if TLab <> nil then TLab.Visible := False ;
            if TLab <> nil then TLab.Caption := '' ;
            end
        else if (j=0) and (TLab <> nil) then TLab.Caption := Titre ;
        {$IFDEF CCS3}
        if i>3 then
           BEGIN
           if TCob<>Nil then TCob.Visible:=False ;
           if TLab<>Nil then TLab.Visible:=False ;
           END ;
       {$ENDIF}
        end
    else if (TypeChp='EDIT') then for j:=0 to Length(Suffixe) do
        begin
        if j=0 then Suf:='' else Suf:=Suffixe[j];
{$IFDEF EAGLCLIENT}
        TEdi:=THEdit (FF.FindComponent (Name+Num+Suf)) ;
{$ELSE}
        TEdi:=THDBEdit (FF.FindComponent (Name+Num+Suf)) ;
{$ENDIF}
        TLab:=THLabel (FF.FindComponent ('T'+Name+Num+Suf)) ;
        if Titre = '' then
            begin
            if TEdi <> nil then TEdi.Visible := False ;
            if TLab <> nil then TLab.Visible := False ;
            if TLab <> nil then TLab.Caption := '' ;
            end
        else if (j=0) and (TLab <> nil) then TLab.Caption := Titre ;
        {$IFDEF CCS3}
        if i>3 then
           BEGIN
           if TEdi<>Nil then TEdi.Visible:=False ;
           if TLab<>Nil then TLab.Visible:=False ;
           END ;
       {$ENDIF}
        end;
    end;
end;

Function GCXXWhereChampLibre (FF:TForm; xx_where:String; TypeChp : String; Name:String; Nbre:Integer; Suffixe:String='') : string;
var Num, Suf, stMCob : String;
    ctrlName, signe : string ;
    TEdi : THEdit;
    TChb : TCheckBox;
    TCob : THValComboBox;
    MCob : THMultiValComboBox ;
    i, j : Integer;
begin
TCob:=Nil;
TEdi:=Nil;
MCob:=Nil ;
For i:=1 to Nbre do
    begin
    if (i < 10) then Num := IntToStr(i) else Num := 'A';

    if (TypeChp='BOOL') then
       begin
       TChb:=TCheckBox(FF.FindComponent(Name+Num)) ;
       if (TChb<>nil) then
          begin
          if (TChb.State=cbChecked) then
             begin
             if xx_where<>'' then xx_where:=xx_where+' AND ' ;
             xx_where:=xx_where+Name+Num+'="X"' ;
             end;
          if (TChb.State=cbUnChecked) then
             begin
             if xx_where<>'' then xx_where:=xx_where+' AND ' ;
             xx_where:=xx_where+Name+Num+'="-"' ;
             end;
          end;
       end ;

    if (TypeChp='DATE') or (typeChp='EDIT') or (typeChp='COMBO') or (typeChp='MULTICOMBO') then
       begin
       for j:=0 to Length(Suffixe) do
           begin
           if j=0 then Suf:='' else Suf:=Suffixe[j];
           ctrlName:=Name+Num+Suf;
           if Length(Suffixe) > 0 then
              begin
              if j=0 then signe:='>='
                     else signe:='<=';
              end
           else if typeChp = 'MULTICOMBO' then signe := ' IN '
           else signe := '=' ;

           if typeChp = 'COMBO' then
             TCob := THValComboBox(FF.FindComponent(ctrlName)) else
           if typeChp = 'MULTICOMBO' then
             MCob := THMultiValComboBox(FF.FindComponent(ctrlName)) else TEdi := THEdit(FF.FindComponent(ctrlName));

           if typeChp='MULTICOMBO' then     // OT 24/02/2003
              begin
              if (MCob <> nil) and (MCob.Value <> '') and (MCob.Value <> TraduireMemoire('<< Tous >>')) then
                 begin
                 if xx_where <> '' then xx_where := xx_where + ' AND ' ;
                 stMCob := MCob.Value ;
                 stMCob := '("' + StringReplace (stMCob, ';', '","', [rfReplaceAll]) ;
                 if Copy (MCob.value, Length (MCob.value), 1) = ';' then
                   stMCob := Copy (stMCob, 1, Length (stMCob) - 2) else
                   stMCob := stMCob + '"' ;
                   if (GetPresentation = ART_ORLI) and (Name = 'GA2_FAMILLENIV') then
                     Num := IntToStr(i + Nbre - 2) ;
                   xx_where := xx_where + Name + Num + signe + stMCob + ')' ;
                 end ;
              end else
           if typeChp='COMBO' then
              begin
              if (TCob<>nil) and (TCob.Value<>'') and (TCob.Value<>TraduireMemoire('<< Tous >>')) then
                 begin
                 if xx_where<>'' then xx_where:=xx_where+' AND ' ;
                 xx_where:=xx_where+Name+Num+signe+'"'+TCob.Value+'"' ;
                 end ;
              end else
           if typeChp='DATE' then
              begin
              if (TEdi<>nil) and (TEdi.Text<>'01/01/1900') and (TEdi.Text<>'31/12/2099') then
                 begin
                 if xx_where<>'' then xx_where:=xx_where+' AND ' ;
                 xx_where:=xx_where+Name+Num+signe+'"'+USDateTime(StrToDate(TEdi.Text))+'"' ;
                 end ;
              end else
           if typeChp='EDIT' then
              begin
              if (TEdi<>nil) and (TEdi.Text<>'') then
                 begin
                 if xx_where<>'' then xx_where:=xx_where+' AND ' ;
                 xx_where:=xx_where+Name+Num+signe+TEdi.Text ;
                 end ;
              end;
           end ;
       end ;
    end;
Result:=xx_where;
end;

{***********A.G.L.***********************************************
Auteur  ...... : JT (eQualité 10819)
Créé le ...... : 29/09/2003
Modifié le ... :
Description .. : Permet de récupérer le titre d'une zone libre fournisseur.
Suite ........ : car ceux sont ceux des clients qui sont dans V_PGI.DECHAMPS
Mots clefs ... : LIBRE;TITRE
*****************************************************************}
Function LibTableLibreCliFour(VenteAchat, NomDuChamp : string) : string;
var CodeLibre : string;
begin
  if VenteAchat = 'ACH' then
    CodeLibre := CodeTableLibre_2(NomDuChamp,'FOU')
    else
    CodeLibre := CodeTableLibre_2(NomDuChamp,'CLI');
  if CodeLibre <> '' then
    Result := RechDom('GCZONELIBRE',CodeLibre,False)
    else
    Result := ChampToLibelle(NomDuChamp);
end;

{***********A.G.L.***********************************************
Auteur  ...... : JCF
Créé le ...... : 08/11/2000
Modifié le ... : 08/11/2000
Description .. : Permet de récupérer le titre d'une zone libre.
Suite ........ : Rend FALSE si la zone n'est pas une zone libre reconnue
Suite ........ : Attention : le titre n'est modifié que si c'est une zone libre.
Mots clefs ... : LIBRE;TITRE
*****************************************************************}
Function GCTitreZoneLibre (Name : String; Var TitreZone : String;AvecPrefixe : boolean=false;  LePrefixe:string=''): Boolean;
Var St,sprefixe : string;
begin
St:=CodeTableLibre_2 (Name);
if St = '' then result := False
else begin
   result := True;
   if AvecPrefixe then
    begin
    if LePrefixe<>'' then sPrefixe:= LePrefixe else
    sPrefixe := copy(name, 1, pos('_', name)-1);

    IF sPrefixe = 'RAC' then sPrefixe := 'RRC' ;
    if Sprefixe = 'AFO' then sPrefixe := 'ALF';
    TitreZone := RechDomZoneLibre (St, False, '', false, sPrefixe);
    end
   else
   TitreZone := RechDomZoneLibre (St, False);
   if Length(TitreZone)> 0 then if (copy(TitreZone,1,2)='.-') then TitreZone := '';
   end;
end;

Function CodeTableLibre_2 (Name : String; CLiFou : string = ''): string;
         // fonction qui compose le code accès table libre en
         // fct nom du champ passé.
var code, nom , nomsuite1: string;
ind : integer;
begin
result:=''; Code := ''; nomsuite1 :='';
ind := Pos ('_',Name);
if ind = 0 then  exit;
nom := Copy (name,1, ind-1);
// PA Eviter de traduire les bornes de fin
if Copy(name,length(name), 1) = '_' then exit;
    // on recherche le type de fiche
if (nom='RPE') then begin
    ind := Pos ('RPETABLELIBRE',name);
    if ind<> 0 then begin
      nomsuite1 := 'RPETABLELIBRE';
      if ((ind + Length(nomsuite1)-1) = Length (name)) then
      begin
       result :='';
       exit;
      end;
      code := 'TL'+ Copy (name, ind+ length (nomsuite1),1);
      result := code;
      exit;
      end;
    end
else if (nom = 'TYTC') or (nom = 'YTC') then begin
        ind := Pos ('FOU',name);
        if (ind <> 0) then
            begin
            code := 'F';
            nomsuite1 := 'TABLELIBREFOU';
            end else
            begin
            code := 'C';
            nomsuite1 := 'TABLELIBRETIERS';
            end;
        end
else if (nom = 'TGCL') or (nom = 'GCL') then code := 'V'
else if (nom = 'TAFF') or (nom = 'AFF') or (nom = 'TEAF') or (nom = 'EAF')then begin
// mcd 23/12/02 else if (nom = 'TAFF') or (nom = 'AFF') then begin
        code := 'M';
        nomsuite1 := 'LIBREAFF';
        end
// Modif BTP
else if (nom = 'TBLO') or (nom = 'BLO') then begin
        code := 'A';
        nomsuite1 := 'LIBREART';
        end
// Fin Modif BTP
else if (nom = 'TGA') or (nom = 'GA') or
        (nom = 'TGL') or (nom = 'GL') or
        (nom = 'TGPF') or (nom = 'GPF') then   //JD
        begin
        code := 'A';
        nomsuite1 := 'LIBREART';
        end
        // mcd 30/01/01
else if (nom = 'TGP') or (nom = 'GP') then begin
        if Pos ('DATE',name)<>0 then begin
              Code :='P';
             end
        else if Pos ('PIECE',name) <> 0 then begin
           // a faire pour appel à bonne table de libelle
           // fait pour champ GP_LIBREPIECE
              Code :='P';
              Nomsuite1:='LIBREPIECE';
              end
        else begin
             ind := Pos ('LIBRETIERS',name);
             if ind<> 0 then begin
             nomsuite1 := 'LIBRETIERS';
             if ((ind + Length(nomsuite1)-1) = Length (name)) then begin
             result :='';
             exit;
             end;
                 // à revoir si table libre fournissuer : CF ...
          if CLiFou = 'FOU' then
            code := 'FT'+ Copy (name, ind+ length (nomsuite1),1)
            else
            code := 'CT'+ Copy (name, ind+ length (nomsuite1),1);
          result := code;
          exit;
          end;
        ind := Pos ('TIERSSAL',name);
        if ind<> 0 then begin
          nomsuite1 := 'TIERSSAL';
          if ((ind + Length(nomsuite1)-1) = Length (name)) then begin
             result :='';
             exit;
             end;
          code := 'CR'+ Copy (name, ind+ length (nomsuite1),1);
          result := code;
          exit;
          end;
        ind := Pos ('AFF',name);
        if ind<> 0 then
          begin
          code := 'M';
          nomsuite1 := 'LIBREAFF';
          end;
          end;
        end
        // mcd 12/02/01 specfi table affaire : aftableaubord
else if (nom = 'TATB') or (nom = 'ATB') then begin
        ind := Pos ('LIBRETIERS',name);
        if ind<> 0 then begin
          nomsuite1 := 'LIBRETIERS';
          if ((ind + Length(nomsuite1)-1) = Length (name)) then begin
             result :='';
             exit;
             end;
                 // à revoir si table libre fournissuer : CF ...
          if CLiFou = 'FOU' then
            code := 'FT'+ Copy (name, ind+ length (nomsuite1),1)
            else
            code := 'CT'+ Copy (name, ind+ length (nomsuite1),1);
          result := code;
          exit;
          end;
        ind := Pos ('AFF',name);
        if ind<> 0 then
          begin
          code := 'M';
          nomsuite1 := 'LIBREAFF';
          end;
        ind := Pos ('LIBRERES',name);
        if ind<> 0 then
          begin
          code := 'R';
          nomsuite1 := 'LIBRERES';
          end;
        ind := Pos ('LIBREART',name);
        if ind<> 0 then
          begin
          code := 'A';
          nomsuite1 := 'LIBREART';
          end;
        end
else if (nom = 'TARS') or (nom = 'ARS') then begin
        code := 'R';
        nomsuite1 := 'LIBRERES';
        end
else if (nom = 'TET') or (nom = 'ET') then
begin
  code := 'E';
  nomsuite1 := 'LIBREET';
end
// modif OT 26/11/2002
else if (nom = 'TGDE') or (nom = 'GDE')then
begin
  code := 'E';
  nomsuite1 := 'LIBREDEP';
end
else if (nom = 'TC') or (nom = 'C') then begin
        code := 'B';
        nomsuite1 := 'LIBRECONTACT';
        end

//planning
else if (nom = 'APL') then begin
        code := 'T';
        nomsuite1 := 'LIBRETACHE';
        end

else if (nom = 'ATA') then begin
        code := 'T';
        nomsuite1 := 'LIBRETACHE';
        end

else if (nom = 'AFM') then begin
        code := 'T';
        nomsuite1 := 'LIBRETACHE';
        end
else if (nom = 'RSC') then
				begin
          if Pos('RSCLIBTABLE',Name) > 0 then
          	Code := 'CL'
          else if Pos('RSCLIBVAL',Name) > 0 then
          	Code := 'VL'
          else if Pos('RSCLIBTEXTE',Name) > 0 then
          	Code := 'TL'
          else if Pos('RSCLIBDATE',Name) > 0 then
          	Code := 'DL'
          else if Pos('RSCLIBBOOL',Name) > 0 then
          	Code := 'BL';
          // du fait du decalage des numeros dna sla grc ...grmffffff
          result := Code + Copy(name,length(Name),1);
        end
else exit;
      // on recherche si combo
ind := Pos (nomsuite1,name);
if (ind <>0) then  Code := code +'T'
else if (Pos ('VALLIBREFOU',name) <> 0) then begin Code := code+'M'; nomsuite1 := 'VALLIBREFOU'; end
else if (Pos ('VALLIBRE',name) <> 0) then begin Code := code+'M'; nomsuite1 := 'VALLIBRE'; end
else if (Pos ('DATELIBREPIECE',name) <> 0) then begin Code := code+'D'; nomsuite1 := 'DATELIBREPIECE'; end
else if (Pos ('DATELIBREFOU',name) <> 0) then begin Code := code+'D'; nomsuite1 := 'DATELIBREFOU'; end
else if (Pos ('DATELIBRE',name) <> 0) then begin Code := code+'D'; nomsuite1 := 'DATELIBRE'; end
else if (Pos ('CHARLIBRE',name) <> 0) then begin Code := code+'C'; nomsuite1 := 'CHARLIBRE'; end
else if (Pos ('TEXTELIBRE',name) <> 0) then begin Code := code+'C'; nomsuite1 := 'TEXTELIBRE'; end
else if (Pos ('BOOLLIBRE',name) <> 0) then begin Code := code+'B'; nomsuite1 := 'BOOLLIBRE'; end
else if (Pos ('RESSOURCE',name) <> 0) then begin Code := code+'R'; nomsuite1 := 'RESSOURCE'; end
else exit;

    // on recherche le dernier caractère
ind :=Pos(nomsuite1,name) ;
if ((ind + Length(nomsuite1)-1) = Length (name)) or (name='ARS_RESSOURCELIE') then begin
        // cas ou aucun nombre derriere, il ne faut pas traiter le champ
        // mcd 01/07/2003 ajour test ressourcelie car efface le libelle
   result :='';
   exit;
   end;
code := code+ Copy (name, ind+ length (nomsuite1),1);
result := code;
end;

{***********A.G.L.***********************************************
Auteur  ...... : JCF
Créé le ...... : 23/09/2000
Modifié le ... : 16/10/2000
Description .. : Modification des libellés des champs pour affichage dans
Suite ........ : les listes.
Suite ........ : Ajout en ifdef des renommages spécifiques aux marchés
Mots clefs ... : LIBELLE;DICO;CHAMPS
*****************************************************************}
Procedure GCModifLibelles;
Var iTable, iChamp : integer;
    NLib : string;
		Mcd : IMCDServiceCOM;
    Table     : ITableCOM ;
    FieldList : IEnumerator ;
  	Field     : IFieldCOMEx ;

  procedure MasqueCube ( Nomchamp : string );
  begin
    MCD := TMCD.GetMcd;
    if not mcd.loaded then mcd.WaitLoaded();
    Field := Mcd.getField(Nomchamp) as IfieldComEx;
    if (Length(Field.Libelle)>0) and (copy(Field.Libelle,1,2)='.-') and (pos( 'C',Field.Control)<>0) then
    Field.Control:= FindEtReplace(Field.Control,'C','',True);
  end;

  procedure GereLibellesLibres (stPrefixe,stTablette : string; isGRC : boolean) ;
  var
    Table     : ITableCOM ;
    NLib : string;
  begin
    MCD := TMCD.GetMcd;
    if not mcd.loaded then mcd.WaitLoaded();
    Table := Mcd.GetTable(Mcd.PrefixeToTable(stPrefixe));
    FieldList := Table.Fields;
    FieldList.Reset();

    While FieldList.MoveNext do
    begin
      Field := FieldList.Current as IFieldCOMEx ;
      if isGRC then NLib:=GRCCodeTableLibre (Field.Name) else NLib := CodeTableLibre_2 (Field.name);
      if NLib <> '' then
      {$IFDEF CCS3}
      if ((stPrefixe='YTC')  and
      	 ((copy(NLib,3,1)='A') or (strToInt(copy(NLib,3,1))>3))) or ((stPrefixe='C')  and
         ((copy(NLib,3,1)='A') or (strToInt(copy(NLib,3,1))>3))) or (stPrefixe='RD1') then
      begin
         Field.Libelle :='.-invisible';
      end else
      {$ENDIF}
      	Field.Libelle :=rechdom (stTablette,NLib,False);
      MasqueCube(Field.Name);
    end;
  end;
begin

  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();

  GereLibellesLibres('YTC','GCZONELIBRE',False);                  // TIERS - zones libres
  GereLibellesLibres('ET','GCZONELIBRE',False);                   // ETABLISS - zones libres
  GereLibellesLibres('C','GCZONELIBRE',False);                    // CONTACTS - zones libres
  GereLibellesLibres('GP','GCZONELIBRE',False);                    // PIECE - zones libres
  GereLibellesLibres('GL','GCZONELIBRE',False);                    // LIGNE - zones libres
  if (ctxAffaire in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte) then
  begin
    GereLibellesLibres('AFF','GCZONELIBRE',False);             // AFFAIRE - zones libres
    GereLibellesLibres('ARS','GCZONELIBRE',False);             // RESSOURCE - zones libres
    GereLibellesLibres('ATA','GCZONELIBRE',False);             // TACHE - zones libres
    GereLibellesLibres('APL','GCZONELIBRE',False);             // PLANNING - zones libres
  end;

  if (ctxGRC in V_PGI.PGIContexte) then       // GRC - zones libres
  begin
    GereLibellesLibres('RPR','RTLIBCHAMPSLIBRES',True);
    GereLibellesLibres('RPC','RTLIBTABLECOMPL',True);
    GereLibellesLibres('RAC','RTLIBCHAMPSLIBRES',True);
    GereLibellesLibres('RAG','RTLIBCHAMPSLIBRES',True);
    GereLibellesLibres('RPE','RTLIBCHAMPSLIBRES',True);
    GereLibellesLibres('RSC','RTLIBCHPLIBSUSPECTS',True);
    GereLibellesLibres('RD1','RTLIBCHAMPS001',True);
    GereLibellesLibres('RD2','RTLIBCHAMPS002',True);
    GereLibellesLibres('RCH','RTLIBCHAINAGES',True);
    GereLibellesLibres('RPA','RTLIBCHAMPSLIBRES',True);
    GereLibellesLibres('RPG','RTLIBCHAINAGES',True);
  end;

  // ARTICLES
  Table:= Mcd.GetTable(Mcd.PrefixeToTable('GA'));
  FieldList := Table.Fields;
  FieldList.Reset();

  While FieldList.MoveNext do
  begin
    Field := FieldList.current as IFieldCOMEx;
    if Field.Name = 'GA_FAMILLENIV1' then Field.Libelle := RechDom('GCLIBFAMILLE','LF1',False)
    else if Field.Name = 'GA_FAMILLENIV2' then Field.Libelle := RechDom('GCLIBFAMILLE','LF2',False)
    else if Field.Name = 'GA_FAMILLENIV3' then Field.Libelle := RechDom('GCLIBFAMILLE','LF3',False)
    else if Field.Name = 'GA_CODEDIM1' then Field.Libelle := RechDom('GCCATEGORIEDIM','DI1',False)
    else if Field.Name = 'GA_CODEDIM2' then Field.Libelle := RechDom('GCCATEGORIEDIM','DI2',False)
    else if Field.Name = 'GA_CODEDIM3' then Field.Libelle := RechDom('GCCATEGORIEDIM','DI3',False)
    else if Field.Name = 'GA_CODEDIM4' then Field.Libelle := RechDom('GCCATEGORIEDIM','DI4',False)
    else if Field.Name = 'GA_CODEDIM5' then Field.Libelle := RechDom('GCCATEGORIEDIM','DI5',False)
    else if Field.Name = 'GA_COLLECTION' then Field.Libelle := RechDom('GCZONELIBRE','AS0',False)
    else if (ctxMode in V_PGI.PGIContexte) and (Field.Name = 'GA_PVHT') then Field.Libelle := 'Prix Négoce (HT)'
    else if (ctxMode in V_PGI.PGIContexte) and (Field.Name = 'GA_PVTTC') then Field.Libelle := 'Prix Détail (TTC)'
    else
    begin
      NLib := CodeTableLibre_2 (Field.Name);
      if NLib <> '' then
      {$IFDEF CCS3}
        if ((copy(NLib,3,1)='A') or (strToInt(copy(NLib,3,1))>3)) then  Field.Libelle :='.-invisible'
        else
      {$ENDIF}
      Field.Libelle :=rechdom ('GCZONELIBRE',NLib,False);
      MasqueCube(Field.name);
    end;
  end;

// ARTICLECOMPL
  Table:= Mcd.GetTable(Mcd.PrefixeToTable('GA2'));
  FieldList := Table.Fields;
  FieldList.Reset();

  While FieldList.MoveNext do
  begin
    Field := FieldList.current as IFieldCOMEx;
    if Field.name = 'GA2_FAMILLENIV4' then Field.Libelle := RechDom('GCLIBFAMILLE','LF4',False)
    else if Field.name = 'GA2_FAMILLENIV5' then Field.Libelle := RechDom('GCLIBFAMILLE','LF5',False)
    else if Field.name = 'GA2_FAMILLENIV6' then Field.Libelle := RechDom('GCLIBFAMILLE','LF6',False)
    else if Field.name = 'GA2_FAMILLENIV7' then Field.Libelle := RechDom('GCLIBFAMILLE','LF7',False)
    else if Field.name = 'GA2_FAMILLENIV8' then Field.Libelle := RechDom('GCLIBFAMILLE','LF8',False)
    else if Field.name = 'GA2_STATART1' then Field.Libelle := RechDom('GCZONELIBRE','AS1',False)
    else if Field.name = 'GA2_STATART2' then Field.Libelle := RechDom('GCZONELIBRE','AS2',False) ;
  end;

// DIMMASQUE
  Table:= Mcd.GetTable(Mcd.PrefixeToTable('GDM'));
  FieldList := Table.Fields;
  FieldList.Reset();

  While FieldList.MoveNext do
  begin
    Field := FieldList.current as IFieldCOMEx;
    if Field.name = 'GDM_TYPE1' then Field.Libelle := RechDom('GCCATEGORIEDIM','DI1',False)
    else if Field.name = 'GDM_TYPE2' then Field.Libelle := RechDom('GCCATEGORIEDIM','DI2',False)
    else if Field.name = 'GDM_TYPE3' then Field.Libelle := RechDom('GCCATEGORIEDIM','DI3',False)
    else if Field.name = 'GDM_TYPE4' then Field.Libelle := RechDom('GCCATEGORIEDIM','DI4',False)
    else if Field.name = 'GDM_TYPE5' then Field.Libelle := RechDom('GCCATEGORIEDIM','DI5',False) ;
  end;

// Si Mono-dépôt, le libellé du champ Dépôt est renommé en Etablissement
	if not VH_GC.GCMultiDepots then
  begin
      // DISPO - dépôt
    Table:= Mcd.GetTable(Mcd.PrefixeToTable('GQ'));
    FieldList := Table.Fields;
    FieldList.Reset();

    While FieldList.MoveNext do
    begin
    	Field := FieldList.current as IFieldCOMEx;
      if (Field.name = 'GQ_DEPOT') then Field.Libelle := 'Etablissement';
    end;

    // LISTEINVENT - dépôt
    Table:= Mcd.GetTable(Mcd.PrefixeToTable('GIE'));
    FieldList := Table.Fields;
    FieldList.Reset();

    While FieldList.MoveNext do
    begin
    	Field := FieldList.current as IFieldCOMEx;
      if (Field.name = 'GIE_DEPOT') then Field.Libelle := 'Etablissement';
    end;

    // TRANSINVENT - dépôt
    Table:= Mcd.GetTable(Mcd.PrefixeToTable('GIT'));
    FieldList := Table.Fields;
    FieldList.Reset();

    While FieldList.MoveNext do
    begin
    	Field := FieldList.current as IFieldCOMEx;
      if (Field.name = 'GIT_DEPOT') then Field.Libelle := 'Etablissement';
    end;
  end;

  // LIGNE - familles - dépôt - zones libres
  Table:= Mcd.GetTable(Mcd.PrefixeToTable('GL'));
  FieldList := Table.Fields;
  FieldList.Reset();

  While FieldList.MoveNext do
  begin
   	Field := FieldList.current as IFieldCOMEx;
    if Field.name = 'GL_FAMILLENIV1' then Field.Libelle := RechDom('GCLIBFAMILLE','LF1',False)
    else if Field.name = 'GL_FAMILLENIV2' then Field.Libelle := RechDom('GCLIBFAMILLE','LF2',False)
    else if Field.name = 'GL_FAMILLENIV3' then Field.Libelle := RechDom('GCLIBFAMILLE','LF3',False)
    else if (not VH_GC.GCMultiDepots) and (Field.name = 'GL_DEPOT') then Field.Libelle := 'Etablissement'
    else
    begin
      NLib := CodeTableLibre_2 (Field.Name);
      if NLib <> '' then Field.Libelle :=rechdom ('GCZONELIBRE',NLib,False);
    end;
  end;
  if VH_GC.GCIfDefCEGID then
  begin
    Table:= Mcd.GetTable(Mcd.PrefixeToTable('T'));
    FieldList := Table.Fields;
    FieldList.Reset();

    While FieldList.MoveNext do
    begin
      Field := FieldList.current as IFieldCOMEx;
      if (Field.name = 'T_RVA') then Field.Libelle := 'Site Web';
    end;
  end;
end;

// GRC No champ to code tablette
Function GRCCodeTableLibre (Name : String): string;
         // fonction qui compose le code accès table libre en
         // fct nom du champ passé.
var code, nom , SuffX, stInd: string;
ind : integer;
begin
result:=''; Code := ''; SuffX :='';
ind := Pos ('_',Name);
if ind = 0 then  exit;
nom := Copy (name,1, ind-1);
if Copy(name,length(name), 1) = '_' then exit;

if nom='RPR' then
  begin
  if (Pos ('RPRLIBTEXTE',name) <> 0)      then begin Code :='TL' ; SuffX:='RPRLIBTEXTE';  end
  else if (Pos ('RPRLIBVAL',name) <> 0)   then begin Code :='VL' ; SuffX:='RPRLIBVAL';   end
  else if (Pos ('RPRLIBTABLE',name) <> 0) then begin Code :='CL' ; SuffX:='RPRLIBTABLE'; end
  else if (Pos ('RPRLIBMUL',name) <> 0)   then begin Code :='ML' ; SuffX:='RPRLIBMUL';   end
  else if (Pos ('RPRLIBBOOL',name) <> 0)  then begin Code :='BL' ; SuffX:='RPRLIBBOOL';  end
  else if (Pos ('RPRLIBDATE',name) <> 0)  then begin Code :='DL' ; SuffX:='RPRLIBDATE';  end
  else if (Pos ('BLOCNOTE',name) <> 0)  then begin Code :='BLO' ; SuffX:='BLOCNOTE';  end
  end;

if (nom='RD1' ) then
  begin
  if (Pos ('RD1LIBTEXTE',name) <> 0)      then begin Code :='TL' ; SuffX:='RD1LIBTEXTE';  end
  else if (Pos ('RD1LIBVAL',name) <> 0)   then begin Code :='VL' ; SuffX:='RD1LIBVAL';   end
  else if (Pos ('RD1LIBTABLE',name) <> 0) then begin Code :='CL' ; SuffX:='RD1LIBTABLE'; end
  else if (Pos ('RD1LIBMUL',name) <> 0)   then begin Code :='ML' ; SuffX:='RD1LIBMUL';   end
  else if (Pos ('RD1LIBBOOL',name) <> 0)  then begin Code :='BL' ; SuffX:='RD1LIBBOOL';  end
  else if (Pos ('RD1LIBDATE',name) <> 0)  then begin Code :='DL' ; SuffX:='RD1LIBDATE';  end
  end;

if (nom='RD2' ) then
  begin
  if (Pos ('RD2LIBTEXTE',name) <> 0)      then begin Code :='TL' ; SuffX:='RD2LIBTEXTE';  end
  else if (Pos ('RD2LIBVAL',name) <> 0)   then begin Code :='VL' ; SuffX:='RD2LIBVAL';   end
  else if (Pos ('RD2LIBTABLE',name) <> 0) then begin Code :='CL' ; SuffX:='RD2LIBTABLE'; end
  else if (Pos ('RD2LIBMUL',name) <> 0)   then begin Code :='ML' ; SuffX:='RD2LIBMUL';   end
  else if (Pos ('RD2LIBBOOL',name) <> 0)  then begin Code :='BL' ; SuffX:='RD2LIBBOOL';  end
  else if (Pos ('RD2LIBDATE',name) <> 0)  then begin Code :='DL' ; SuffX:='RD2LIBDATE';  end
  end;

if (nom='RSC') then
  begin
  if (Pos ('RSCLIBTEXTE',name) <> 0)      then begin Code :='TL' ; SuffX:='RSCLIBTEXTE';  end
  else if (Pos ('RSCLIBVAL',name) <> 0)   then begin Code :='VL' ; SuffX:='RSCLIBVAL';   end
  else if (Pos ('RSCLIBTABLE',name) <> 0) then begin Code :='CL' ; SuffX:='RSCLIBTABLE'; end
  else if (Pos ('RSCLIBMUL',name) <> 0)   then begin Code :='ML' ; SuffX:='RSCLIBMUL';   end
  else if (Pos ('RSCLIBBOOL',name) <> 0)  then begin Code :='BL' ; SuffX:='RSCLIBBOOL';  end
  else if (Pos ('RSCLIBDATE',name) <> 0)  then begin Code :='DL' ; SuffX:='RSCLIBDATE';  end
  end;

if nom='RPC' then
  begin
  if (Pos ('RPCLIBTEXTE',name) <> 0)      then begin Code :='TX' ; SuffX:='RPCLIBTEXTE';  end
  else if (Pos ('RPCLIBVAL',name) <> 0)   then begin Code :='VL' ; SuffX:='RPCLIBVAL';   end
  else if (Pos ('RPCLIBTABLE',name) <> 0) then begin Code :='TL' ; SuffX:='RPCLIBTABLE'; end
  else if (Pos ('RPCLIBBOOL',name) <> 0)  then begin Code :='BL' ; SuffX:='RPCLIBBOOL';  end
  else if (Pos ('RPCLIBDATE',name) <> 0)  then begin Code :='DL' ; SuffX:='RPCLIBDATE';  end
  end;

if (nom='RAC') or (nom='RAG') or (nom='RPA') then
  begin
  if (Pos ('TABLELIBRE',name) <> 0)      then begin Code :='AL' ; SuffX:='TABLELIBRE';  end
  end;

if nom = 'RPE' then
  begin
  if (Pos ('TABLELIBREPER',name) <> 0)   then begin Code :='PL' ; SuffX:='TABLELIBREPER';  end
  end;
if (nom = 'RCH') or (nom='RPG') then
  begin
  if (Pos ('TABLELIBRECH',name) <> 0)   then begin Code :='RH' ; SuffX:='TABLELIBRECH';  end
  end;

// recherche de l'indice
if code<>'' then
  begin
  if SuffX = 'BLOCNOTE' then
    begin
    result := code;
    end
  else
    begin
    ind :=Pos(SuffX,name) ;
    stInd:= Copy (name, ind+length (SuffX),length(name)-length(SuffX));
    if length(stInd)>1 then  stInd:=chr(ord('A')+ strtoint(stInd)-10) ;
    result := code+ stInd;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : JCF
Créé le ...... : 09/04/2001
Modifié le ... : 09/04/2001
Description .. : Ajout de l'établissement dans la barre de bas d'écran
Mots clefs ... : ETABLISSEMENT;BARRE
*****************************************************************}
Procedure GCTripoteStatus ;
Var St : String ;
BEGIN
if ctxMode in V_PGI.PGIContexte then
  BEGIN
  if GetParamSoc('SO_ETABLISDEFAUT') <> '' then
    begin
    // QLoc:=OpenSql('Select ET_LIBELLE From ETABLISS Where ET_ETABLISSEMENT="'+GetParamSoc('SO_ETABLISDEFAUT')+'"',True) ;
    // StSeule:=QLoc.Fields[0].AsString ; Ferme(QLoc) ;
    St := GetDefStatus + '               ' + TraduireMemoire ('Etablissement : ') +
          GetParamSoc('SO_ETABLISDEFAUT') + ' - ' +
          RechDom('TTETABLISSEMENT',GetParamSoc('SO_ETABLISDEFAUT'),FALSE);
    ChgStatus(St) ;
    end;
  END;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : JCF
Créé le ...... : 23/09/2000
Modifié le ... : 16/10/2000
Description .. : Modification de la condition Where de la tablette
Suite ........ : GCZONELIBRE relative aux titres des zones libres.
Mots clefs ... : CHOIXCOD;ZONE;LIBRE;COMBO
*****************************************************************}
Procedure GCModifTitreZonesLibres (Prefixe : String; stCode : string = 'ZLI') ;
Var ii : integer ;
    St : String;
BEGIN
if stCode = 'ZLA' then st := st + 'GCZONELIBREART'
else if stcode = 'ZLT' then st := st + 'GCZONELIBRETIE'
else if stcode = 'ZLP' then st := st + 'GCZONELIBRESAV'
else if stcode = 'ZLC' then st := st + 'GCZONELIBRECON'
else if stcode = 'RRC' then st := st + 'RTLIBACTIONS'
else if stcode = 'RCI' then st := st + 'RTLIBCIBLAGE'
else if stcode = 'RRP' then st := st + 'RTLIBPROJET'
else if stcode = 'RRO' then st := st + 'RTLIBOPERATION'
else if stcode = 'RRS' then st := st + 'RTLIBPERSPECTIVE'
else if stcode = 'RRC' then st := st + 'RTLIBACTIONS'
else if stcode = 'RCI' then st := st + 'RTLIBCIBLAGE'
else
  St:='GCZONELIBRE' ;
ii:=TTToNum(St) ;
if ii>0 then
   begin
   if Prefixe='' then V_PGI.DECombos[ii].Where:= StWhereZoneLibre else
      begin
      StWhereZoneLibre:=V_PGI.DECombos[ii].Where;
      V_PGI.DECombos[ii].Where:='CC_TYPE="' + stCode + '" AND CC_CODE > "'+Prefixe+'" AND CC_CODE < "'+Prefixe+'Z"' ;
{$IFDEF CCS3}
      V_PGI.DECombos[ii].Where:=V_PGI.DECombos[ii].Where+' AND CC_CODE < "'+Prefixe+'4"' ;
{$ENDIF}
      end;
   end;
END ;


Function ChangeLibelleCombo (t,s : string) : string ;
var  st,stTablette,stCode : string;
begin
Result:=s;
st:=s;
if ReadTokenSt(St) <> '&#@' then Exit;
stTablette:=ReadTokenSt(St);
stCode:=ReadTokenSt(St);
if (stCode='') or (stTablette='') then exit;
Result:=rechdom (stTablette,stCode,False);
end;

function ChangeBoolLibre ( NomChamp : string ; FF : Tform; SansUnderScore : boolean=false ): boolean;
var ST2, NomChampLib : string;
begin
result:=true; {visible}
// PL le 23/05/07 : on peut choisir de ne pas tenir compte du _ au début du nom du champ
// pour retrouver son libellé libre
NomChampLib := NomChamp;
if SansUnderScore then
  if (NomChampLib[1]='_') then
    NomChampLib := copy(NomChampLib, 2, length(NomChampLib)-1);

st2:=ChampToLibelle(NomChampLib);
TOMTOFSetControlCaption(FF,NomChamp,st2);
if (copy(st2,1,2)='.-') then begin TOMTOFSetControlVisible(FF,NomChamp,false); result:=False; end;
end;

function ChangeBoolLibre2 ( NomChamp,Nomzone  : string ; FF : Tform; SansUnderScore : boolean=false ): boolean;
var ST2, NomChampLib : string;
begin
result:=true; {visible}
// PL le 23/05/07 : on peut choisir de ne pas tenir compte du _ au début du nom du champ
// pour retrouver son libellé libre
NomChampLib := Nomzone;
if SansUnderScore then
  if (NomChampLib[1]='_') then
    NomChampLib := copy(NomChampLib, 2, length(NomChampLib)-1);

st2:=ChampToLibelle(NomChampLib);
TOMTOFSetControlCaption(FF,NomChamp,st2);
if (copy(st2,1,2)='.-') then begin TOMTOFSetControlVisible(FF,NomChamp,false); result:=False; end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Paul Chapuis
Créé le ...... : 02/12/2002
Modifié le ... : 02/12/2002
Description .. : positionne le caption du controle dont le nom est passé en
Suite ........ : argument en fonction :
Suite ........ :    soit du libellé du champ associé (FocusControl)
Suite ........ :    soit du libellé du champ dont le nom est extrait du nom du
Suite ........ : controle qui doit alors être sous la forme Tnom_du_champ :
Suite ........ :            exemple : TRAC_TABLELIBRE1
Mots clefs ... : LIBRE;LIBELLE
*****************************************************************}
function ChangeLibre2 ( NomChamp : string ; FF : Tform ) : boolean;
var CC : THLabel;
begin
result:=True; {Visible}
CC:=THLabel(TomTofGetControl(FF,NomChamp));            
if CC=Nil then Exit;
if CC.FocusControl <> nil then CC.Caption:=ChampToLibelle(CC.FocusControl.Name);
if (CC.Caption='??') or (CC.FocusControl = nil) then CC.Caption:=ChampToLibelle(copy(NomChamp,2,length(Nomchamp)));
if (Length(CC.Caption)>0) and (copy(CC.Caption,1,2)='.-') then
   begin
   CC.visible:=false;
   TomTofSetControlVisible(FF,CC.Name+'_', False);
   TomTofSetControlVisible(FF,CC.FocusControl.Name, False);
   TomTofSetControlVisible(FF,CC.FocusControl.Name+'_', False);
   Result:=False;
   end;
end;

Procedure MajChampsLibresArticle(FF: Tform; stPrefixe:string='GA';MasqueSiVide:boolean=true;TSTablesLibres:string='PTABLESLIBRES';TSZonesLibres : string='PZONESLIBRES' );
var i         : integer;
    AuMoinsUn : boolean;
begin

  AuMoinsUn:=(not MasqueSiVide);

  For i:=1 to 10 do
    if i<10 then
      AuMoinsUn:=(ChangeLibre2('T'+stPrefixe+'_LIBREART'+intToStr(i),FF) or AuMoinsUn)
    else
      AuMoinsUn:=(ChangeLibre2('T'+stPrefixe+'_LIBREARTA',FF)or AuMoinsUn);

  TOMTOFSetControlVisible(FF,TSTablesLibres, AuMoinsUn) ; {Invisible si aucun controle}

  if TSTablesLibres<>TSZonesLibres then AuMoinsUn:=MasqueSiVide;

  For i:=1 to 3 do AuMoinsUn:=(ChangeBoolLibre(stPrefixe+'_BOOLLIBRE'+intToStr(i),FF)or AuMoinsUn);
  For i:=1 to 3 do AuMoinsUn:=(ChangeLibre2('T'+stPrefixe+'_VALLIBRE'+intToStr(i),FF)or AuMoinsUn);
  For i:=1 to 3 do AuMoinsUn:=(ChangeLibre2('T'+stPrefixe+'_DATELIBRE'+intToStr(i),FF)or AuMoinsUn);

  {$IFDEF CCS3}
  TOMTOFSetControlVisible(FF,TSZonesLibres,False) ;
  {$ELSE}
  TOMTOFSetControlVisible(FF,TSZonesLibres,AuMoinsUn) ;
  {$ENDIF}

end;



Function RTPhonemeSearch(St1 : String) : String ;
begin
result:=Phoneme(st1);
end;
{$IFDEF AUCASOU}

{ Version de Phoneme sans suppression du dernier caractère si S,T,D etc pour construire l'argument
  du LIKE. sinon recherche sur bres devient LIKE "BRE%" et retourne bressuire (ok) et bretagne (!)
}

Function SansMrMme(St1 : string) : string ;
var St2 : string;
BEGIN
If Copy(St1,1,3)='MR.' Then St2:=Copy(St1,4,100) else
If Copy(St1,1,6)='MADAME' Then St2:=Copy(St1,7,100) else
If Copy(St1,1,8)='MONSIEUR' Then St2:=Copy(St1,9,100) else
if Copy(St1,1,2)='MR' Then St2:=Copy(St1,3,100) else
if Copy(St1,1,3)='DR.' Then St2:=Copy(St1,4,100) else
if Copy(St1,1,3)='DR ' Then St2:=Copy(St1,4,100) else
if Copy(St1,1,4)='MMR.' Then St2:=Copy(St1,5,100)else
if Copy(St1,1,4)='MMR ' Then St2:=Copy(St1,5,100)else
if Copy(St1,1,4)='MLE.' Then St2:=Copy(St1,5,100)else
if Copy(St1,1,4)='MLE ' Then St2:=Copy(St1,5,100)  else
if Copy(St1,1,6)='MELLE.' Then St2:=Copy(St1,7,100)  else
if Copy(St1,1,6)='MELLE ' Then St2:=Copy(St1,7,100) else
if Copy(St1,1,5)='MLLE.' Then St2:=Copy(St1,6,100) else
if Copy(St1,1,4)='MLLE' Then St2:=Copy(St1,5,100) else
if Copy(St1,1,4)='MME.' Then St2:=Copy(St1,5,100) else
if Copy(St1,1,3)='MME' Then St2:=Copy(St1,4,100) else
if Copy(St1,1,2)='M.' Then St2:=Copy(St1,3,100) else
if Copy(St1,1,2)='M ' Then St2:=Copy(St1,3,100) Else St2:=St1 ;
SansMrMme:=St2 ;
End ;


Const MaxCouple = 80 ;
Couple : Array [1..MaxCouple,1..2] of string[10] =(
   ('.',''),('-',''),('''',''),(',',''),(';',''),('CT','KT'),('QH','K'),('GH','G'),('NG ','N'),('ZZ','Z'),
   ('NH','N'),('Ç','S'),(' H',' '),('DD','D'),('AIN ','IN '),
   ('W','V'),('BB','B'),('OUIN','OIN'),('CHR','KR'),('FF','F'),
   ('OUX','OU'),('PH','F'),('CE','SE'),('CCI','CHI'),('CI','SI'),('CO','KO'),
   ('KH','K'),('LH','L'),('CA','KA'),('CU','KU'),('AI','E'),('AY','E'),
   ('EY','E'),('EAUX ','O'),('RH','R'),('EAU','O'),('QU','K'),('NN','N'),
   ('MM','M'),('SS','S'),('ILL','Y'),('AM','EM'),
   ('EI','E'),('CS ','X '),
   ('TH','T'),('SCH','CH'),('GNI','NI'),('GN','NI'),('AUX ','O '),
   ('AU','O'),('EUX','E'),('EU','E'),('EE','E'),('PP','P'),('MPT','NT'),
   ('NT ','N '),('CK','K'),('Y ','I '),('IER ','IE '),('RR','R'),('TT','T'),
   ('OLL','OL'),('ELL','EL'),('EAN','AN'),('PT','T'),('ALL','AL'),('ULL','UL'),
   ('CR','KR'),('Y','I'),('MB','NB'),('IX','I'),('TIA','SIA'),('RG ','R '),
   ('RC ','R '),('EZ ','ES '),('GE','JE'),('GI','JI'),('C ','K '),('Z','S'),
   (' ',''));

{dans HENT1
   ('.',''),('-',''),('''',''),(',',''),(';',''),('CT','KT'),('QH','K'),('GH','G'),('NG ','N'),('ZZ','Z'),
   ('NH','N'),('Ç','S'),(' H',' '),('DD','D'),('AIN ','IN '),
   ('W','V'),('LT ',' '),('BB','B'),('OUIN','OIN'),('CHR','KR'),('FF','F'),
   ('OUX','OU'),('PH','F'),('CE','SE'),('CCI','CHI'),('CI','SI'),('CO','KO'),
   ('KH','K'),('LH','L'),('CA','KA'),('CU','KU'),('AI','E'),('AY','E'),
   ('EY','E'),('EAUX ','O'),('RH','R'),('EAU','O'),('QU','K'),('NN','N'),
   ('MM','M'),('SS','S'),('ILL','Y'),('AM','EM'),
   ('EI','E'),('D ',' '),('T ',' '),('CS ','X '),
   ('S ',' '),('TH','T'),('SCH','CH'),('GNI','NI'),('GN','NI'),('AUX ','O '),
   ('AU','O'),('EUX','E'),('EU','E'),('EE','E'),('PP','P'),('MPT','NT'),
   ('NT ','N '),('CK','K'),('Y ','I '),('IER ','IE '),('RR','R'),('TT','T'),
   ('OLL','OL'),('ELL','EL'),('EAN','AN'),('PT','T'),('ALL','AL'),('ULL','UL'),
   ('CR','KR'),('Y','I'),('MB','NB'),('IX','I'),('TIA','SIA'),('RG ','R '),
   ('RC ','R '),('EZ ','ES '),('GE','JE'),('GI','JI'),('C ','K '),('Z','S'),
   (' ',''));
}
Var St2 : String ;
    i,Pos1 : Integer ;

BEGIN
St1:=UpperCase(trim(St1)) ;
St1:=SansMrMme(St1) ;
If Copy(St1,2,2)='. ' Then  St1:=Copy(St1,4,100)+' '+Copy(St1,1,2) ;
If Copy(St1,2,1)='.' Then St1:=Copy(St1,3,100)+' '+Copy(St1,1,2) ;
St2:=' '+UpperCase(trim(St1))+' ' ;
For i:=1 to MaxCouple do
   BEGIN
   Pos1:=Pos(Couple[i,1],St2) ;
   If Pos1>0 Then St2:=Copy(St2,1,Pos1-1)+Couple[i,2]+Copy(St2,Pos1+Length(Couple[i,1]),100) ;
   Pos1:=Pos(Couple[i,1],St2) ;
   If Pos1>0 Then St2:=Copy(St2,1,Pos1-1)+Couple[i,2]+Copy(St2,Pos1+Length(Couple[i,1]),100) ;
   END ;
RTPhonemeSearch:=trim(St2) ;
End ;
{$ENDIF}

Procedure MajChampsLibresGED( FF : Tform ) ;
var i: integer;
begin
if FF=Nil then  exit;
For i:=1 to 3 do ChangeLibre2('TRTD_TABLELIBREGED'+intToStr(i),FF);
end;

Procedure AGLGCModifTitreZonesLibres ( Parms : array of variant ; nb : integer ) ;
BEGIN
if Nb > 1 then GCModifTitreZonesLibres (Parms [0],Parms[1])
					else GCModifTitreZonesLibres (Parms [0]) ;
END ;

Procedure AGLEditMonarch ( Parms : array of variant ; nb : integer ) ;
BEGIN
EditMonarch (Parms [0]) ;
END ;

procedure AGLMajChampsLibresArticle( parms: array of variant; nb: integer ) ;
begin
  MajChampsLibresArticle(TForm(Longint(Parms[0])),string(Parms[1]),boolean(Parms[2]),string(Parms[3]),string(Parms[4])) ;
end;

procedure GCTestDepotOblig;
begin
  if (VH_GC.GCDepotDefaut = '') and
     (not ((CtxScot in V_PGI.PGIContexte) or (CtxTempo in V_PGI.PGIContexte) or (CtxChr in V_PGI.PGIContexte))) then
    HShowMessage('0;ATTENTION;Paramétrage société incomplet. Veuillez exécuter l''utilitaire d''affection des dépôts et renseigner un dépôt par défaut.;W;O;O;O;','','');
end;

{***********A.G.L.***********************************************
Auteur  ...... : JS
Créé le ...... : 22/04/2004
Modifié le ... :   /  /
Description .. : Mise à jour de la table journal d'événement
Mots clefs ... : JOURNAL;EVENEMENT;
*****************************************************************}
function MAJJnalEvent(TypeEvt, Etat, Libelle, BlocNote : string; FileName: string ='') : integer;
{$IFNDEF EAGLCLIENT}
var TobJournal : TOB;
    TSqlJournal : TQuery;
{$ENDIF}
begin
  Result := 0;
{$IFDEF EAGLCLIENT}
{$ELSE EAGLCLIENT}
  TobJournal := TOB.Create('JNALEVENT', Nil, -1) ;
  try
    TobJournal.PutValue('GEV_TYPEEVENT',TypeEvt);
    TobJournal.PutValue('GEV_ETATEVENT',Etat);
    TobJournal.PutValue('GEV_LIBELLE',TraduireMemoire(Libelle));
    TobJournal.PutValue('GEV_DATEEVENT',Date);
    TobJournal.PutValue('GEV_UTILISATEUR',V_PGI.User);
    TOBJournal.PutValue('GEV_BLOCNOTE', TraduireMemoire(BlocNote));
    TSqlJournal := OpenSQL ('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT', True);
    try
      if Not TSqlJournal.EOF then Result := TSqlJournal.Fields[0].AsInteger;
    finally
      Ferme (TSqlJournal);
    end;
    Inc (Result);
    TOBJournal.PutValue ('GEV_NUMEVENT', Result);
    BEGINTRANS;
    TRY
      TOBJournal.InsertDB (Nil) ;
      //
      if FileName <> '' then StockeDocumentLie (TfafEvt,TobJournal,FileName);
      COMMITTRANS;
    EXCEPT
      ROLLBACK;
    END;
  finally
    TobJournal.Free;
  end;
{$ENDIF EAGLCLIENT}
end;

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 07/09/2004
Modifié le ... :
Description .. : Transformation d'un NomChamp in("xx","yy") en NomChamp="xx" or NomChamp="yy"
Mots clefs ... : TIERS
*****************************************************************}
function TransformeLesInToOr(stIn : string) : string;
var UppIn, NomChamp, LeRetour : string;
    PosChar : integer;
begin
  Result := StIn;
  if stIn = '' then exit;
  UppIn := Uppercase(stIn);
  Result := UppIn;
  if pos(' IN',UppIn) = 0 then exit;
  PosChar := pos(' IN',UppIn);
  NomChamp := copy(UppIn,1,PosChar-1);
  UppIn := copy(UppIn,Pos('"',UppIn),length(UppIn));
  UppIn := copy(UppIn,1,Pos(')',UppIn)-1);
  LeRetour := '';
  while UppIn<>'' do
    LeRetour := LeRetour + ' OR (' + NomChamp + '=' + ReadTokenPipe(UppIn,',') + ')';
  LeRetour := copy(LeRetour,4,length(LeRetour));
  Result := LeRetour;
end;


function GereCommercial : boolean;
begin
  Result := GetParamSocSecur ('SO_GERECOMMERCIAL',False);
end;

{***********A.G.L.***********************************************
Auteur  ...... : JTR - eQualité 10204
Créé le ...... : 27/09/2005
Modifié le ... :   /  /
Description .. : Permet de traiter des champs sur un TForm
Suite ........ : Etablissement, Dépot et Domaine : affecte la valeur de l'utilisateur courant
Suite ........ : Représentant : cache si on ne gère pas les représentants
Mots clefs ... :
*****************************************************************}
procedure TraiteParametresSurTForm(FF : TForm; LesChamps : string=''; AffecteFiltre : boolean=false);
var Cpt, Cpt1 : integer;
    LeControl : TControl;
    LeChamp, stTexte : string;
    TabChamps : array of string;
    MaxChamps : integer;
    OkFiltre : boolean;

    procedure TraiteEtablissement;
    begin
      if not (Lecontrol is THLabel) then
        PositionneEtabUser(LeControl, False, True, True);
    end;

    procedure TraiteDepot;
    begin
      if not (Lecontrol is THLabel) then
        PositionneDepotUser(LeControl, False, True, True);
    end;

    procedure TraiteDomaine;
    var TabDomaine : array of variant;
    begin
      if not (Lecontrol is THLabel) then
      begin
        SetLength(TabDomaine, 2);
        TabDomaine[0] := Longint(FF);
        TabDomaine[1] := LeControl.Name;
        AGLPositionneDomaineUser(TabDomaine, 2);
      end;
    end;

    procedure TraiteRepresentant;
    begin
      {$IFDEF GCGC}
      LeControl.Visible := GereCommercial;
      {$ELSE GCGC}
      LeControl.Visible := true;
      {$ENDIF GCGC}
    end;

begin
  if FF = nil then exit;
  if pos('ETABLISSEMENT', LesChamps) = 0 then
    LesChamps := LesChamps + ';ETABLISSEMENT';
  if pos('DEPOT', LesChamps) = 0 then
    LesChamps := LesChamps + ';DEPOT';
  if pos('DOMAINE', LesChamps) = 0 then
    LesChamps := LesChamps + ';DOMAINE';
  if pos('REPRESENTANT', LesChamps) = 0 then
    LesChamps := LesChamps + ';REPRESENTANT';
  if copy(LesChamps, 1, 1) = ';' then
    LesChamps := Copy(LesChamps, 2, length(LesChamps));  
  MaxChamps := NbOccurenceString(LesChamps, ';') + 1;
  SetLength(TabChamps, MaxChamps);
  Cpt := 0;
  while LesChamps <> '' do
  begin
    TabChamps[Cpt] := ReadTokenSt(LesChamps);
    Cpt := Cpt + 1;
  end;
  for Cpt := 0 to FF.ComponentCount - 1 do
  begin
    OkFiltre := True;
    LeControl := TControl(FF.Components[Cpt]);
    for Cpt1 := 0 to high(TabChamps) do
    begin
      LeChamp := TabChamps[Cpt1];
      if pos(LeChamp, LeControl.name) > 0 then
      begin
        if (pos('ETABLISSEMENT', LeChamp) > 0) then TraiteEtablissement
        else if (pos('DEPOT', LeChamp) > 0) then TraiteDepot
        else if (pos('DOMAINE', LeChamp) > 0) then TraiteDomaine
        else if (pos('REPRESENTANT', LeChamp) > 0) then TraiteRepresentant
        else
        begin
          OkFiltre := False;
          Continue;
        end;
      end;
    end;
    if (OkFiltre) and (AffecteFiltre) then
    begin
      stTexte := '';
      LeControl.Tag := 0;
      if LeControl is THValComboBox then
        stTexte := Trim(THValComboBox(LeControl).Text)
      else if LeControl is THMultiValComboBox then
        stTexte := Trim(THMultiValComboBox(LeControl).Text);
      if (stTexte <> '') and (not LeControl.Enabled) then
        LeControl.Tag := -9970;
    end;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 18/11/2004
Modifié le ... :
Description .. : Compte le nombre d'occurences dans une chaîne
Mots clefs ... :
*****************************************************************}
function NbOccurenceString(Chaine, Occurence : string) : integer;
var Cpt, LgOcc : integer;
begin
  result := 0;
  if Chaine = '' then exit;
  LgOcc := length(Occurence);
  Cpt := 1;
  while Cpt <= length(Chaine) do
  begin
    if Copy(Chaine,Cpt,LgOcc) = Occurence then
    begin
      Cpt := Cpt + LgOcc;
      Result := Result + 1;
    end else
      Cpt := Cpt + 1;
  end;
end;

Procedure RTCreerProspectPourClient ;
var
  TOBProspects: TOB;
  Q : TQuery;
begin
  if PGIAsk('Voulez-vous initialiser les infos complémentaires pour les tiers clients/prospects ?','')<>mrYes then exit ;
  TOBProspects:=Tob.create('PROSPECTS',Nil,-1);
  {$IFDEF GIGI}
  Q:=OPENSQL('select t_auxiliaire from tiers where '+VH_GC.AfTypTiersGRCGI+' and NOT EXISTS (select rpr_auxiliaire from prospects where rpr_auxiliaire=tiers.t_auxiliaire)',true);
  {$ELSE GIGI}
  Q:=OPENSQL('select t_auxiliaire from tiers where (t_natureauxi="CLI" or t_natureauxi="PRO") and NOT EXISTS (select rpr_auxiliaire from prospects where rpr_auxiliaire=tiers.t_auxiliaire)',true);
  {$ENDIF GIGI}
  InitMove(Q.recordcount, '');
  try
    while not Q.EOF do
    begin
      TOBProspects.InitValeurs(False);
      TOBProspects.PutValue('RPR_AUXILIAIRE',Q.Fields[0].AsString);
      if not TOBProspects.InsertOrUpdateDB(False) then
        V_PGI.IoError:=oeUnknown ;
      Q.Next ;
      MoveCur(false);
    end;
  finally
    FiniMove();
    Ferme(Q);
    TOBProspects.free;
  end;
  PGIInfo('Traitement de mise à jour terminé', 'Infos complémentaires');
end;


Procedure RTAlimCleTelephonTiers ;
var
  Notel,Cletel : string;
  Q : TQuery;
begin
  if PGIAsk('Voulez-vous mettre à jour la clé téléphone ?', 'MAJ Clé téléphone') <> mrYes then exit
  else
  begin
    Q := OpenSQL ('SELECT t_tiers,t_telephone,t_cletelephone FROM tiers where t_telephone<>""',True) ;
    try
      while Not Q.EOF do
      begin
        Cletel:='';
        Notel := Q.FindField('t_telephone').AsString;
        Cletel:=CleTelephone(Notel);
        if Cletel <> '' then
          ExecuteSQL('UPDATE tiers SET t_cletelephone="'+Cletel+'" where t_tiers="'+Q.FindField('t_tiers').asString+'"');
        Q.Next;
      end;
    finally
      Ferme(Q) ;
    end;  
    PGIInfo('Traitement de mise à jour terminé', 'MAJ Clé téléphone');
  end;
end;

Procedure RTAlimTelephonContact ;
var
  Notel,Cletel : string;
  Q : TQuery;
begin
  if PGIAsk('Voulez-vous initialiser les no téléphones contacts ?', 'MAJ no téléphone') <> mrYes then exit
  else
  begin
    Q := OpenSQL ('SELECT C_AUXILIAIRE,C_NUMEROCONTACT,c_telephone FROM contact where c_telephone<>"" and c_typecontact="T"',True) ;
    try
      while Not Q.EOF do
      begin
        Cletel:='';
        Notel := Q.FindField('c_telephone').AsString;
        Cletel:=CleTelephone(Notel);
        if Cletel <> '' then
          ExecuteSQL('UPDATE contact SET c_cletelephone="'+Cletel+'" where C_TYPECONTACT="T" and C_NUMEROCONTACT='+
                    Q.FindField('C_NUMEROCONTACT').asString+' and C_AUXILIAIRE="'+Q.FindField('C_AUXILIAIRE').asString+'"');
        Q.Next;
      end;
    finally
      Ferme(Q);
    end;
    PGIInfo('Traitement de mise à jour terminé', 'MAJ Clé téléphone');
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Pierre Lenormand
Créé le ...... : 24/05/07
Modifié le ... :
Description .. : renomme les champs AAA_BBBBBB en __AAA_BBBBBB
                  et les champs _AAA_BBBBBB en AAA_BBBBBB
                  utilisé quand on joue sur plusieurs présentations d'un même écran
                  la fiche Suspects par exemple
Suite ........ : argument en fonction :
Suite ........ :    nom champ sans indice, indice dédut de l'intervalle et indice fin de l'intervalle
Suite ........ :    Ecran
                    RendInvisible : rend invisible le champ __AAA_BBBBBB et visible le champ AAA_BBBBBB
                    une fois l'intervertion faite
Suite ........ : dans le résultat : le nb de champs effectivement modifiés
Suite ........ :
Mots clefs ... : LIBRE;CHAMP
*****************************************************************}
function IntervertirNomChampsXX ( FF:Tform; sChampSansIndice:string; iDebut, iFin:integer; RendInvisible:boolean=false) : integer;
var
ii : integer;
sNomCourant : string;
TCtrl, TCtrl_, TCtrlxx : TControl;
begin
result := -1;
if (sChampSansIndice = '') or (iDebut < 0) or (iFin < 0) or (iDebut > iFin) then
  exit;

result := 0;
for ii:=iDebut to iFin do
  begin
    sNomCourant := sChampSansIndice + inttostr(ii);
    TCtrl := TControl(TomTofGetControl(FF, sNomCourant));
    TCtrl_ := TControl(TomTofGetControl(FF, sNomCourant+'_'));
    TCtrlxx := TControl(TomTofGetControl(FF, '_'+sNomCourant));
    if (TCtrlxx = nil) then
      continue;


    // AAA_BBBBBB => __AAA_BBBBBB
    if (TCtrl <> nil) then
      TCtrl.name := '__' + sNomCourant;
    if (TCtrl_ <> nil) then
      TCtrl_.name := '__' + sNomCourant+'_';
    // _AAA_BBBBBB => AAA_BBBBBB
    TCtrlxx.name := sNomCourant;

    if RendInvisible then
      begin
        TomTofSetControlVisible(FF, sNomCourant, true);
        if (TCtrl <> nil) then
          TomTofSetControlVisible(FF, '__' + sNomCourant, False);
        if (TCtrl_ <> nil) then
          TomTofSetControlVisible(FF, '__' + sNomCourant+'_', False);
      end;

    Inc (result);
  end;
end;

Function RTTypeproduitCRM : string ;
begin
Result := '';
{$IFDEF COMFI}
if GetParamSocSecur ('SO_COMFI',False) = True then Result := 'COMFI';
{$ENDIF COMFI}
end;



Initialization
RegisterAglProc('MajChampsLibresArticle',True,4,AGLMajChampsLibresArticle) ;
RegisterAglProc('GCModifTitreZonesLibres',False,2,AGLGCModifTitreZonesLibres) ;
RegisterAglProc('EditMonarch', False, 1, AGLEditMonarch) ;
{$ENDIF}

end.
