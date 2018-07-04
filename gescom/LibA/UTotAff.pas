unit UTOTAff;

interface
uses Classes,UTOT,SysUtils,Dicobtp,HCtrls,
{$IFDEF EAGLCLIENT}

{$ELSE}
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
 HEnt1,UtilPaieAffaire,ActiviteUtil;

Type
    TOT_AFTTYPERESSOURCE = Class (TOT)
        Function GetTitre : Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        procedure OnUpdateRecord  ; override ;
        end;
    TOT_AFFAIREPART1 = Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_AFFAIREPART2 = Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_AFFAIREPART3 = Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_AFTLIBREAFF1 = Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_AFTLIBREAFF2 = Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_AFTLIBREAFF3 = Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_AFTLIBREAFF4 = Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_AFTLIBREAFF5 = Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_AFTLIBREAFF6 = Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_AFTLIBREAFF7 = Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_AFTLIBREAFF8 = Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_AFTLIBREAFF9 = Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_AFTLIBREAFFA = Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_AFTLIBRERES1 = Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_AFTLIBRERES2 = Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_AFTLIBRERES3= Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_AFTLIBRERES4= Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_AFTLIBRERES5= Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_AFTLIBRERES6= Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_AFTLIBRERES7= Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_AFTLIBRERES8= Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_AFTLIBRERES9= Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_AFTLIBRERESA= Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_AFTTYPEHEURE = Class (TOT)
        procedure OnUpdateRecord  ; override ;
        end;
    TOT_AFETATAFFAIRE = Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        procedure OnUpdateRecord  ; override ;
        end;
    TOT_AFTACTIVITEORIGINE = Class (TOT)
        procedure OnDeleteRecord  ; override ;
        procedure OnUpdateRecord  ; override ;
        end;
    TOT_AFTREGROUPEFACT = Class (TOT)
        procedure OnDeleteRecord  ; override ;
        procedure OnUpdateRecord  ; override ;
        end;
    TOT_AFDEPARTEMENT = Class (TOT)
        procedure OnDeleteRecord  ; override ;
        end;
    TOT_AFCOMPTAAFFAIRE = Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        end;
    TOT_AFTNIVEAUDIPLOME = Class (TOT)
        procedure OnDeleteRecord  ; override ;
        end;
    TOT_AFFAMILLETACHE = Class (TOT)
        procedure OnDeleteRecord  ; override ;
        end;
    TOT_AFLIBRETACHE1 = Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        end;
    TOT_AFLIBRETACHE2 = Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        end;
    TOT_AFLIBRETACHE3 = Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        end;
    TOT_AFLIBRETACHE4 = Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        end;
    TOT_AFLIBRETACHE5 = Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        end;
    TOT_AFLIBRETACHE6 = Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        end;
    TOT_AFLIBRETACHE7 = Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        end;
    TOT_AFLIBRETACHE8 = Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        end;
    TOT_AFLIBRETACHE9 = Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        end;
    TOT_AFLIBRETACHEA = Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        end;
    TOT_AFTSTANDCALEN = Class (TOT)
        Function GetTitre :Hstring; override ;
        procedure OnDeleteRecord  ; override ;
        End;
    TOT_AFTLIENAFFTIERS = Class (TOT)
        Function GetTitre :Hstring; override ;
        end;
    TOT_AFTRESILAFF = Class (TOT)
        Function GetTitre :Hstring; override ;
        end;
    TOT_AFTTARIFRESSOURCE = Class (TOT)
        Function GetTitre :Hstring; override ;
        end;


implementation
//***************** Departement ****************************************
                    // mcd 16/04/02 ajout blocage ...
procedure TOT_AFDEPARTEMENT.OnDeleteRecord;
var SQL: string;
BEGIN
  SQL:='SELECT AFF_AFFAIRE FROM AFFAIRE where AFF_DEPARTEMENT="'+GetField('YX_CODE')+'"' ;
  if ExisteSQL (SQL) then begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier Affaire','');
    Lasterror:=1;
    exit;
    End;
  SQL:='SELECT ARS_RESSOURCE FROM RESSOURCE where ARS_DEPARTEMENT="'+GetField('YX_CODE')+'"' ;
  if ExisteSQL (SQL) then begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier ressource','');
    Lasterror:=1;
    exit;
    End;
End;

 //***************** code comptable affaire ****************************************
                    // mcd 16/04/02 ajout blocage ...
procedure TOT_AFCOMPTAAFFAIRE.OnDeleteRecord;
var SQL: string;
BEGIN
  Sql:='SELECT AFF_AFFAIRE FROM AFFAIRE where AFF_COMPTAAFFAIRE="'+GetField('CC_CODE')+'"' ;
  if ExisteSQL(SQL) then   begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier Affaire','');
    Lasterror:=1;
    exit;
    End;
End;
Function TOT_AFCOMPTAAFFAIRE.GetTitre : Hstring;
BEGIN
result:=TraduitGA('Famille comptable affaire');
End;

//***************** Niveau diplôme ****************************************
                    // mcd 16/04/02 ajout blocage ...
procedure TOT_AFTNIVEAUDIPLOME.OnDeleteRecord;
var SQL: string;
BEGIN
  SQL:='SELECT ACR_SALARIE FROM COMPETRESSOURCE where ACR_NIVEAUDIPLOME="'+GetField('CC_CODE')+'"' ;
  if ExisteSQL(SQL) then   begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier Compétence/ressource','');
    Lasterror:=1;
    exit;
    End;
End;


//***************** Types de ressources ****************************************
Function TOT_AFTTYPERESSOURCE.GetTitre : Hstring;
BEGIN
result:=TraduitGa ('Type des ressources');
End;
procedure TOT_AFTTYPERESSOURCE.OnDeleteRecord;
var code,lib,stit: string;
    CodeUse : boolean;
BEGIN
  Code := GetField('CC_CODE');
  Lib :=  RechDom('AFTTYPERESSOURCE',Code,False);
  stit := format('Type de ressource : %s - %s',[Code,Lib]);
    // controle des liens avec  la table Affaire  et Profilgener
  CodeUse := SupTablesLiees ('RESSOURCE', 'ARS_TYPERESSOURCE', Code, '' , false);

// Gestion des codes obligatoires
if (Code = 'ST') or (Code = 'SAL') then
   Begin
   PGIBoxAF('Suppression interdite, ce type de ressource est obligatoire',stit);
   Lasterror:=1; Exit;
   End;
{$IFDEF BTP}
if (Code = 'INT') or (Code = 'LOC')or (Code = 'MAT') then
    Begin
    PGIBoxAF('Suppression interdite, ce type de ressource est obligatoire',stit);
    Lasterror:=1;
    exit;
    End;
{$ENDIF}
// Gestion des codes utilisés
if CodeUse then
   Begin
   PGIBoxAF('Suppression interdite, ce type de ressource est utilisé',stit);
   Lasterror:=1; Exit;
   End;
End;

procedure TOT_AFTTYPERESSOURCE.OnUpdateRecord;
Var Code,Lib,Stit : String;
BEGIN
Code := GetField('CC_CODE');
Lib :=  RechDom('AFTTYPERESSOURCE',Code,False);
stit := format('Type de ressource : %s - %s',[Code,Lib]);
if (Code = 'ST') then PGIBoxAF('Attention ce code référence obligatoirement des sous-traitants',stit);
if (Code = 'SAL') then PGIBoxAF('Attention ce code référence obligatoirement des salariés',stit);
{$IFDEF BTP}
if (Code = 'INT') then PGIBoxAF('Attention ce code référence obligatoirement des intérimaires',stit);
if (Code = 'LOC') then PGIBoxAF('Attention ce code référence obligatoirement des matériels loués',stit);
if (Code = 'MAT') then PGIBoxAF('Attention ce code référence obligatoirement des matériels internes',stit);
{$ENDIF}
END;


//************************ Etats de l'affaire **********************************
Function TOT_AFETATAFFAIRE.GetTitre : Hstring;
BEGIN
result:=TraduitGa ('Etat des Affaires');
End;
procedure TOT_AFETATAFFAIRE.OnDeleteRecord;
var code,lib,stit: string;
    CodeUse : boolean;
BEGIN
Code := GetField('CC_CODE'); Lib :=  RechDom('AFETATAFFAIRE',Code,False);
stit := format('Etat de l''affaire : %s - %s',[Code,Lib]);
// controle des liens avec  la table Affaire
CodeUse := SupTablesLiees ('AFFAIRE', 'AFF_ETATAFFAIRE', Code, '' , false);
// Gestion des codes obligatoires
{$ifdef BTP}
if (Code = 'ACP') or (Code = 'ENC') or (Code='TER')  then    
   Begin PGIBoxAF('Suppression interdite, cet état est obligatoire',stit); Lasterror:=1; exit; End;
{$else}
if (Code = 'REF') or (Code = 'ENC') or (Code='ACC') or (Code='CLO') then   //mcd 07/01/03 ajout CLO : des traitement sont fait sur ce code, destrcution interdite
   Begin PGIBoxAF('Suppression interdite, cet état est obligatoire',stit); Lasterror:=1; exit; End;
{$endif}
   // Gestion des codes utilisés
if CodeUse then
   Begin PGIBoxAF('Suppression interdite, cet état est utilisé',stit); Lasterror:=1; exit; End;
End;

procedure TOT_AFETATAFFAIRE.OnUpdateRecord;
Var Code,Lib,Stit : String;
BEGIN
Code := GetField('CC_CODE'); Lib :=  RechDom('AFETATAFFAIRE',Code,False);
stit := format('Etat de l''affaire : %s - %s',[Code,Lib]);
if (Code = 'ACC') then PGIBoxAF('Attention ce code référence obligatoirement des propositions acceptées',stit) else
if (Code = 'REF') then PGIBoxAF('Attention ce code référence obligatoirement des propositions refusées',stit) else
if (Code = 'ENC') then PGIBoxAF('Attention ce code référence obligatoirement des affaires en cours',stit)
else SetField('CC_LIBRE','AFF');

END;
       // tablette ActiviteOrigine
procedure TOT_AFTACTIVITEORIGINE.OnDeleteRecord;
var code,lib,stit: string;
    CodeUse : boolean;
BEGIN
Code := GetField('CC_CODE'); Lib :=  RechDom('AFACTIVITEORIGINE',Code,False);
stit := format('Origine Activité : %s - %s',[Code,Lib]);
// controle des liens avec  la table Affaire
CodeUse := SupTablesLiees ('ACTIVITE', 'ACT_ACTORIGINE', Code, '' , false);
If not CodeUse then CodeUse := SupTablesLiees ('EACTIVITE', 'EAC_ACTORIGINE', Code, '' , false);

// Gestion des codes obligatoires
if (Code = 'ACH') or (Code = 'EAC') or (Code='PLA')or (Code='REP')or (Code='SDE') then
   Begin PGIBoxAF('Suppression interdite, ce code est obligatoire',stit); Lasterror:=1; exit; End;
// Gestion des codes utilisés
if CodeUse then
   Begin PGIBoxAF('Suppression interdite, ce code est utilisé',stit); Lasterror:=1; exit; End;
End;

procedure TOT_AFTACTIVITEORIGINE.OnUpdateRecord;
Var Code,Lib,Stit : String;
BEGIN
Code := GetField('CC_CODE'); Lib :=  RechDom('AFTACTIVITEORIGINE',Code,False);
stit := format('Origine Activité : %s - %s',[Code,Lib]);
if (Code = 'ACH') or (Code = 'EAC') or (Code = 'PLA') or (Code = 'REP') or (Code = 'SDE')
   then PGIBoxAF('Attention ce code est obligatoire',stit) ;

END;

//***************** Regroupement de facturation sur affaires *******************
procedure TOT_AFTREGROUPEFACT.OnDeleteRecord;
var code,lib,stit: string;
    CodeUse : boolean;
BEGIN
  Code := GetField('CC_CODE');
  Lib :=  RechDom('AFTREGROUPEFACT',Code,False);
  stit := format('Regroupement : %s - %s',[Code,Lib]);
    // controle des liens avec  la table Affaire  et Profilgener
  CodeUse := SupTablesLiees ('AFFAIRE', 'AFF_REGROUPEFACT', Code, '' , false);

// Gestion des codes obligatoires
if (Code = 'AUC') or (Code = 'ARE') or (Code = 'AVC') then
   Begin
   PGIBoxAF('Suppression interdite, ce code de regroupement est obligatoire',stit);
   Lasterror:=1; Exit;
   End;
// Gestion des codes utilisés
if CodeUse then
   Begin
   PGIBoxAF('Suppression interdite, ce code de regroupement est utilisé',stit);
   Lasterror:=1; Exit;
   End;
End;

procedure TOT_AFTREGROUPEFACT.OnUpdateRecord;
Var Code,Lib,Stit : String;
BEGIN
Code := GetField('CC_CODE');
Lib :=  RechDom('AFTREGROUPEFACT',Code,False);
stit := format('Regroupement : %s - %s',[Code,Lib]);
if (Code = 'AUC') or (Code = 'ARE') then
   begin
   PGIBoxAF('Attention Code non modifiable',stit);
   Lasterror:=1; Exit;
   end
else
   SetField('CC_LIBRE', 'B');
END;

//********************** Stats affaires ****************************************
// bien laisser les fct Gettitre, sinon, le libellé perso de la table libre n'est pas visible sur equere de saisie
Function TOT_AFTLIBREAFF1.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','MT1',False)
End;

Function TOT_AFTLIBREAFF2.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','MT2',False)
End;

Function TOT_AFTLIBREAFF3.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','MT3',False)
End;

Function TOT_AFTLIBREAFF4.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','MT4',False)
End;

Function TOT_AFTLIBREAFF5.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','MT5',False)
End;

Function TOT_AFTLIBREAFF6.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','MT6',False)
End;

Function TOT_AFTLIBREAFF7.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','MT7',False)
End;

Function TOT_AFTLIBREAFF8.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','MT8',False)
End;

Function TOT_AFTLIBREAFF9.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','MT9',False)
End;

Function TOT_AFTLIBREAFFA.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','MTA',False)
End;

Procedure TOT_AFTLIBREAFF1.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT AFF_LIBREAFF1 from AFFAIRE where AFF_LIBREAFF1="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier','Table Libre');
    Lasterror:=1;
    end ;
SQL:='SELECT GP_LIBREAFF1 from PIECE where GP_LIBREAFF1="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier','Table Libre');
    Lasterror:=1;
    end ;
End;

Procedure TOT_AFTLIBREAFF2.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT AFF_LIBREAFF2 from AFFAIRE where AFF_LIBREAFF2="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier','Table Libre');
    Lasterror:=1;
    end ;
SQL:='SELECT GP_LIBREAFF2 from PIECE where GP_LIBREAFF2="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier','Table Libre');
    Lasterror:=1;
    end ;
End;

Procedure TOT_AFTLIBREAFF3.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT AFF_LIBREAFF3 from AFFAIRE where AFF_LIBREAFF3="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier','Table Libre');
    Lasterror:=1;
    end ;
SQL:='SELECT GP_LIBREAFF3 from PIECE where GP_LIBREAFF3="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier','Table Libre');
    Lasterror:=1;
    end ;
End;

Procedure TOT_AFTLIBREAFF4.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT AFF_LIBREAFF4 from AFFAIRE where AFF_LIBREAFF4="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier','Table Libre');
    Lasterror:=1;
    end ;
End;

Procedure TOT_AFTLIBREAFF5.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT AFF_LIBREAFF5 from AFFAIRE where AFF_LIBREAFF5="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier','Table Libre');
    Lasterror:=1;
    end ;
End;
Procedure TOT_AFTLIBREAFF6.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT AFF_LIBREAFF6 from AFFAIRE where AFF_LIBREAFF6="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier','Table Libre');
    Lasterror:=1;
    end ;
End;
Procedure TOT_AFTLIBREAFF7.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT AFF_LIBREAFF7 from AFFAIRE where AFF_LIBREAFF7="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier','Table Libre');
    Lasterror:=1;
    end ;
End;
Procedure TOT_AFTLIBREAFF8.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT AFF_LIBREAFF8 from AFFAIRE where AFF_LIBREAFF8="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier','Table Libre');
    Lasterror:=1;
    end ;
End;
Procedure TOT_AFTLIBREAFF9.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT AFF_LIBREAFF9 from AFFAIRE where AFF_LIBREAFF9="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier','Table Libre');
    Lasterror:=1;
    end ;
End;
Procedure TOT_AFTLIBREAFFA.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT AFF_LIBREAFFA from AFFAIRE where AFF_LIBREAFFA="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier','Table Libre');
    Lasterror:=1;
    end ;
End;

Function TOT_AFFAIREPART1.GetTitre : Hstring;
BEGIN
result:=TraduitGa ('Partie 1 du code affaire');
End;

procedure TOT_AFFAIREPART1.OnDeleteRecord;
var SQL: string;
BEGIN
  SQL:='SELECT AFF_AFFAIRE FROM AFFAIRE where AFF_AFFAIRE1="'+GetField('CC_CODE')+'"';
  if ExisteSQL(SQL) then   begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier','Partie 1');
    Lasterror:=1;
    exit;
    End;
End;

Function TOT_AFFAIREPART2.GetTitre : Hstring;
BEGIN
result:=TraduitGa ('Partie 2 du code affaire');
End;

procedure TOT_AFFAIREPART2.OnDeleteRecord;
var SQL: string;
BEGIN
  SQL:='SELECT AFF_AFFAIRE FROM AFFAIRE where AFF_AFFAIRE2="'+GetField('CC_CODE')+'"' ;
  if ExisteSQL(SQL) then   begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier','Partie 2');
    Lasterror:=1;
    exit;
    End;
End;

Function TOT_AFFAIREPART3.GetTitre : Hstring;
BEGIN
result:=TraduitGa ('Partie 3 du code affaire');
End;

procedure TOT_AFFAIREPART3.OnDeleteRecord;
var SQL: string;
BEGIN
  SQL:='SELECT AFF_AFFAIRE FROM AFFAIRE where AFF_AFFAIRE3="'+GetField('CC_CODE')+'"' ;
  if ExisteSQL(SQL) then   begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier','Partie 3');
    Lasterror:=1;
    exit;
    End;
End;

Function TOT_AFTLIBRERES1.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','RT1',False)
End;

Function TOT_AFTLIBRERES2.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','RT2',False)
End;

Function TOT_AFTLIBRERES3.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','RT3',False)
End;

Function TOT_AFTLIBRERES4.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','RT4',False)
End;

Function TOT_AFTLIBRERES5.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','RT5',False)
End;

Function TOT_AFTLIBRERES6.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','RT6',False)
End;

Function TOT_AFTLIBRERES7.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','RT7',False)
End;

Function TOT_AFTLIBRERES8.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','RT8',False)
End;

Function TOT_AFTLIBRERES9.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','RT9',False)
End;

Function TOT_AFTLIBRERESA.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','RTA',False)
End;

Procedure TOT_AFTLIBRERES1.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT ARS_LIBRERES1 from RESSOURCE where ARS_LIBRERES1="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier','Table Libre');
    Lasterror:=1;
    end ;
End;
Procedure TOT_AFTLIBRERES2.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT ARS_LIBRERES2 from RESSOURCE where ARS_LIBRERES2="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier','Table Libre');
    Lasterror:=1;
    end ;
End;
Procedure TOT_AFTLIBRERES3.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT ARS_LIBRERES3 from RESSOURCE where ARS_LIBRERES3="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier','Table Libre');
    Lasterror:=1;
    end ;
End;
Procedure TOT_AFTLIBRERES4.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT ARS_LIBRERES4 from RESSOURCE where ARS_LIBRERES4="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier','Table Libre');
    Lasterror:=1;
    end ;
End;
Procedure TOT_AFTLIBRERES5.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT ARS_LIBRERES5 from RESSOURCE where ARS_LIBRERES5="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier','Table Libre');
    Lasterror:=1;
    end ;
End;
Procedure TOT_AFTLIBRERES6.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT ARS_LIBRERES6 from RESSOURCE where ARS_LIBRERES6="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier','Table Libre');
    Lasterror:=1;
    end ;
End;
Procedure TOT_AFTLIBRERES7.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT ARS_LIBRERES7 from RESSOURCE where ARS_LIBRERES7="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier','Table Libre');
    Lasterror:=1;
    end ;
End;
Procedure TOT_AFTLIBRERES8.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT ARS_LIBRERES8 from RESSOURCE where ARS_LIBRERES8="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier','Table Libre');
    Lasterror:=1;
    end ;
End;
Procedure TOT_AFTLIBRERES9.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT ARS_LIBRERES9 from RESSOURCE where ARS_LIBRERES9="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier','Table Libre');
    Lasterror:=1;
    end ;
End;
Procedure TOT_AFTLIBRERESA.OnDeleteRecord;
var  SQL : string ;
begin
// Test de l'existence du code dans les tables
SQL:='SELECT ARS_LIBRERESA from RESSOURCE where ARS_LIBRERESA="'+getField('YX_CODE')+'"' ;
if ExisteSQL (SQL) then begin
    PGIBoxAF('Suppression interdite, ce code existe en fichier','Table Libre');
    Lasterror:=1;
    end ;
End;


procedure TOT_AFTTYPEHEURE.OnUpdateRecord;
Var stTaux : String;
    Tx1,Tx2 : Double;
BEGIN
stTaux := GetField('CC_ABREGE');
if stTaux = '' then PGIBoxAF('Attention le taux de majoration des heures n''est pas valide',TTToLibelle('AFTTYPEHEURE'))
else
   BEGIN
   if Not(GetTauxHSup(stTaux ,Tx1,Tx2)) then
      BEGIN
      PGIBoxAF('Attention le taux de majoration des heures n''est pas valide',TTToLibelle('AFTTYPEHEURE'));
      Lasterror:=1; Exit;
      END;
   END;
END;

procedure TOT_AFFAMILLETACHE.OnDeleteRecord;
var
  SQL : string;
begin
  // Test de l'existence du code dans les tables
  SQL:='SELECT ATA_FAMILLETACHE from TACHE where ATA_FAMILLETACHE="'+getField('CC_CODE')+'"' ;
  if ExisteSQL(SQL) then
    begin
      PGIBoxAF('Suppression interdite, ce code existe en fichier',TTToLibelle('AFFAMILLETACHE'));
      Lasterror:=1;
    end;
end;

Function TOT_AFLIBRETACHE1.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','TT1',False)
End;

Function TOT_AFLIBRETACHE2.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','TT2',False)
End;

Function TOT_AFLIBRETACHE3.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','TT3',False)
End;

Function TOT_AFLIBRETACHE4.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','TT4',False)
End;

Function TOT_AFLIBRETACHE5.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','TT5',False)
End;

Function TOT_AFLIBRETACHE6.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','TT6',False)
End;

Function TOT_AFLIBRETACHE7.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','TT7',False)
End;

Function TOT_AFLIBRETACHE8.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','TT8',False)
End;

Function TOT_AFLIBRETACHE9.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','TT9',False)
End;

Function TOT_AFLIBRETACHEA.GetTitre : Hstring;
BEGIN
result:=RechDom('GCZONELIBRE','TTA',False)
End;

procedure TOT_AFLIBRETACHE1.OnDeleteRecord;
var
  vStSQL1 : string;
  vStSQL2 : string;
begin
  // Test de l'existence du code dans les tables
  vStSQL1 := 'SELECT ATA_LIBRETACHE1 from TACHE where ATA_LIBRETACHE1="'+getField('YX_CODE')+'"' ;
  vStSQL2 := 'SELECT APL_LIBRETACHE1 from AFPLANNING where APL_LIBRETACHE1="'+getField('YX_CODE')+'"' ;
  if ExisteSQL(vStSQL1) then
    begin
      PGIBoxAF('Suppression interdite, ce code existe en fichier',TTToLibelle('LIBRETACHE1'));
      Lasterror:=1;
    end
  else if ExisteSQL(vStSQL2) then
    begin
      PGIBoxAF('Suppression interdite, ce code existe en fichier',TTToLibelle('LIBRETACHE1'));
      Lasterror:=1;
    end
end;

procedure TOT_AFLIBRETACHE2.OnDeleteRecord;
var
  vStSQL1 : string;
  vStSQL2 : string;
begin
  // Test de l'existence du code dans les tables
  vStSQL1 := 'SELECT ATA_LIBRETACHE2 from TACHE where ATA_LIBRETACHE2="'+getField('YX_CODE')+'"' ;
  vStSQL2 := 'SELECT APL_LIBRETACHE2 from AFPLANNING where APL_LIBRETACHE2="'+getField('YX_CODE')+'"' ;
  if ExisteSQL(vStSQL1) then
    begin
      PGIBoxAF('Suppression interdite, ce code existe en fichier',TTToLibelle('LIBRETACHE2'));
      Lasterror:=1;
    end
  else if ExisteSQL(vStSQL2) then
    begin
      PGIBoxAF('Suppression interdite, ce code existe en fichier',TTToLibelle('LIBRETACHE2'));
      Lasterror:=1;
    end
end;

procedure TOT_AFLIBRETACHE3.OnDeleteRecord;
var
  vStSQL1 : string;
  vStSQL2 : string;
begin
  // Test de l'existence du code dans les tables
  vStSQL1 := 'SELECT ATA_LIBRETACHE3 from TACHE where ATA_LIBRETACHE3="'+getField('YX_CODE')+'"' ;
  vStSQL2 := 'SELECT APL_LIBRETACHE3 from AFPLANNING where APL_LIBRETACHE3="'+getField('YX_CODE')+'"' ;
  if ExisteSQL(vStSQL1) then
    begin
      PGIBoxAF('Suppression interdite, ce code existe en fichier',TTToLibelle('LIBRETACHE3'));
      Lasterror:=1;
    end
  else if ExisteSQL(vStSQL2) then
    begin
      PGIBoxAF('Suppression interdite, ce code existe en fichier',TTToLibelle('LIBRETACHE3'));
      Lasterror:=1;
    end
end;

procedure TOT_AFLIBRETACHE4.OnDeleteRecord;
var
  vStSQL1 : string;
  vStSQL2 : string;
begin
  // Test de l'existence du code dans les tables
  vStSQL1 := 'SELECT ATA_LIBRETACHE4 from TACHE where ATA_LIBRETACHE4="'+getField('YX_CODE')+'"' ;
  vStSQL2 := 'SELECT APL_LIBRETACHE4 from AFPLANNING where APL_LIBRETACHE4="'+getField('YX_CODE')+'"' ;
  if ExisteSQL(vStSQL1) then
    begin
      PGIBoxAF('Suppression interdite, ce code existe en fichier',TTToLibelle('LIBRETACHE4'));
      Lasterror:=1;
    end
  else if ExisteSQL(vStSQL2) then
    begin
      PGIBoxAF('Suppression interdite, ce code existe en fichier',TTToLibelle('LIBRETACHE4'));
      Lasterror:=1;
    end
end;

procedure TOT_AFLIBRETACHE5.OnDeleteRecord;
var
  vStSQL1 : string;
  vStSQL2 : string;
begin
  // Test de l'existence du code dans les tables
  vStSQL1 := 'SELECT ATA_LIBRETACHE5 from TACHE where ATA_LIBRETACHE5="'+getField('YX_CODE')+'"' ;
  vStSQL2 := 'SELECT APL_LIBRETACHE5 from AFPLANNING where APL_LIBRETACHE5="'+getField('YX_CODE')+'"' ;
  if ExisteSQL(vStSQL1) then
    begin
      PGIBoxAF('Suppression interdite, ce code existe en fichier',TTToLibelle('LIBRETACHE5'));
      Lasterror:=1;
    end
  else if ExisteSQL(vStSQL2) then
    begin
      PGIBoxAF('Suppression interdite, ce code existe en fichier',TTToLibelle('LIBRETACHE5'));
      Lasterror:=1;
    end
end;

procedure TOT_AFLIBRETACHE6.OnDeleteRecord;
var
  vStSQL1 : string;
  vStSQL2 : string;
begin
  // Test de l'existence du code dans les tables
  vStSQL1 := 'SELECT ATA_LIBRETACHE6 from TACHE where ATA_LIBRETACHE6="'+getField('YX_CODE')+'"' ;
  vStSQL2 := 'SELECT APL_LIBRETACHE6 from AFPLANNING where APL_LIBRETACHE6="'+getField('YX_CODE')+'"' ;
  if ExisteSQL(vStSQL1) then
    begin
      PGIBoxAF('Suppression interdite, ce code existe en fichier',TTToLibelle('LIBRETACHE6'));
      Lasterror:=1;
    end
  else if ExisteSQL(vStSQL2) then
    begin
      PGIBoxAF('Suppression interdite, ce code existe en fichier',TTToLibelle('LIBRETACHE6'));
      Lasterror:=1;
    end
end;

procedure TOT_AFLIBRETACHE7.OnDeleteRecord;
var
  vStSQL1 : string;
  vStSQL2 : string;
begin
  // Test de l'existence du code dans les tables
  vStSQL1 := 'SELECT ATA_LIBRETACHE7 from TACHE where ATA_LIBRETACHE7="'+getField('YX_CODE')+'"' ;
  vStSQL2 := 'SELECT APL_LIBRETACHE7 from AFPLANNING where APL_LIBRETACHE7="'+getField('YX_CODE')+'"' ;
  if ExisteSQL(vStSQL1) then
    begin
      PGIBoxAF('Suppression interdite, ce code existe en fichier',TTToLibelle('LIBRETACHE7'));
      Lasterror:=1;
    end
  else if ExisteSQL(vStSQL2) then
    begin
      PGIBoxAF('Suppression interdite, ce code existe en fichier',TTToLibelle('LIBRETACHE7'));
      Lasterror:=1;
    end
end;

procedure TOT_AFLIBRETACHE8.OnDeleteRecord;
var
  vStSQL1 : string;
  vStSQL2 : string;
begin
  // Test de l'existence du code dans les tables
  vStSQL1 := 'SELECT ATA_LIBRETACHE8 from TACHE where ATA_LIBRETACHE8="'+getField('YX_CODE')+'"' ;
  vStSQL2 := 'SELECT APL_LIBRETACHE8 from AFPLANNING where APL_LIBRETACHE8="'+getField('YX_CODE')+'"' ;
  if ExisteSQL(vStSQL1) then
    begin
      PGIBoxAF('Suppression interdite, ce code existe en fichier',TTToLibelle('LIBRETACHE8'));
      Lasterror:=1;
    end
  else if ExisteSQL(vStSQL2) then
    begin
      PGIBoxAF('Suppression interdite, ce code existe en fichier',TTToLibelle('LIBRETACHE8'));
      Lasterror:=1;
    end
end;

procedure TOT_AFLIBRETACHE9.OnDeleteRecord;
var
  vStSQL1 : string;
  vStSQL2 : string;
begin
  // Test de l'existence du code dans les tables
  vStSQL1 := 'SELECT ATA_LIBRETACHE9 from TACHE where ATA_LIBRETACHE9="'+getField('YX_CODE')+'"' ;
  vStSQL2 := 'SELECT APL_LIBRETACHE9 from AFPLANNING where APL_LIBRETACHE9="'+getField('YX_CODE')+'"' ;
  if ExisteSQL(vStSQL1) then
    begin
      PGIBoxAF('Suppression interdite, ce code existe en fichier',TTToLibelle('LIBRETACHE9'));
      Lasterror:=1;
    end
  else if ExisteSQL(vStSQL2) then
    begin
      PGIBoxAF('Suppression interdite, ce code existe en fichier',TTToLibelle('LIBRETACHE9'));
      Lasterror:=1;
    end
end;

procedure TOT_AFLIBRETACHEA.OnDeleteRecord;
var
  vStSQL1 : string;
  vStSQL2 : string;
begin
  // Test de l'existence du code dans les tables
  vStSQL1 := 'SELECT ATA_LIBRETACHEA from TACHE where ATA_LIBRETACHEA="'+getField('YX_CODE')+'"' ;
  vStSQL2 := 'SELECT APL_LIBRETACHEA from AFPLANNING where APL_LIBRETACHEA="'+getField('YX_CODE')+'"' ;
  if ExisteSQL(vStSQL1) then
    begin
      PGIBoxAF('Suppression interdite, ce code existe en fichier',TTToLibelle('LIBRETACHEA'));
      Lasterror:=1;
    end
  else if ExisteSQL(vStSQL2) then
    begin
      PGIBoxAF('Suppression interdite, ce code existe en fichier',TTToLibelle('LIBRETACHEA'));
      Lasterror:=1;
    end
end;

procedure TOT_AFTSTANDCALEN.OnDeleteRecord;
var
  vStSQL1 : string;
  vStSQL2 : string;
begin
  // Test de l'existence du code dans les tables
  vStSQL1 := 'SELECT ACG_STANDCALEN from CALENDRIERREGLE where ACG_STANDCALEN="'+getField('CC_CODE')+'"' ;
  vStSQL2 := 'SELECT ACA_STANDCALEN from CALENDRIER where ACA_STANDCALEN="'+getField('CC_CODE')+'"' ;
  if ExisteSQL(vStSQL1) then
    begin
      PGIBoxAF('Suppression interdite, ce code existe en fichier Règle de calendrier','Type Calendrier');
      Lasterror:=1;
    end
  else if ExisteSQL(vStSQL2) then
    begin
      PGIBoxAF('Suppression interdite, ce code existe en fichier Calendrier','Type Calendrier');
      Lasterror:=1;
    end
end;

Function TOT_AFTSTANDCALEN.GetTitre : Hstring;
BEGIN
result:=TraduitGA('Type de calendrier (ressource)');
End;





Function TOT_AFTLIENAFFTIERS.GetTitre : Hstring;
BEGIN
result:=TraduitGA('Lien entre les tiers et affaires');
End;

Function TOT_AFTRESILAFF.GetTitre : Hstring;
BEGIN
result:=TraduitGA('Motif résiliation sur affaire');
End;
Function TOT_AFTTARIFRESSOURCE.GetTitre : Hstring;
BEGIN
result:=TraduitGA('Code tarifaire de la ressource');
End;


Initialization
  registerclasses([TOT_AFTTYPERESSOURCE,TOT_AFTTYPEHEURE,TOT_AFFAIREPART1,TOT_AFETATAFFAIRE,TOT_AFTACTIVITEORIGINE]) ;
  registerclasses([TOT_AFTLIBREAFF1,TOT_AFTLIBREAFF2,TOT_AFTLIBREAFF3,TOT_AFTLIBREAFF4,TOT_AFTLIBREAFF5,TOT_AFTLIBREAFF6,TOT_AFTLIBREAFF7,TOT_AFTLIBREAFF8,TOT_AFTLIBREAFF9,TOT_AFTLIBREAFFA]) ;
  registerclasses([TOT_AFTLIBRERES1,TOT_AFTLIBRERES2,TOT_AFTLIBRERES3,TOT_AFTLIBRERES4,TOT_AFTLIBRERES5,TOT_AFTLIBRERES6,TOT_AFTLIBRERES7,TOT_AFTLIBRERES8,TOT_AFTLIBRERES9,TOT_AFTLIBRERESA]) ;
  registerclasses([TOT_AFTREGROUPEFACT,TOT_AFDEPARTEMENT,TOT_AFCOMPTAAFFAIRE,TOT_AFTNIVEAUDIPLOME]);
  registerclasses([TOT_AFFAMILLETACHE, TOT_AFLIBRETACHE1, TOT_AFLIBRETACHE2, TOT_AFLIBRETACHE3, TOT_AFLIBRETACHE4]);
  registerclasses([TOT_AFLIBRETACHE5, TOT_AFLIBRETACHE6, TOT_AFLIBRETACHE7, TOT_AFLIBRETACHE8]);
  registerclasses([TOT_AFLIBRETACHE9, TOT_AFLIBRETACHEA, TOT_AFTSTANDCALEN]);
  registerclasses([ TOT_AFTTYPERESSOURCE, TOT_AFTTARIFRESSOURCE]);
  registerclasses([TOT_AFTRESILAFF, TOT_AFTLIENAFFTIERS]);
end.
