{***********UNITE*************************************************
Auteur  ...... : JP
Créé le ...... : 23/04/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFARTICLE_SELECT ()
Mots clefs ... : TOF;AFARTICLE_SELECT
*****************************************************************}
Unit utofafarticle_select ;

Interface

Uses StdCtrls, 
     Controls,
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
     dbtables, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     HMsgBox,
     UTOF,UTOFAFTRADUCCHAMPLIBRE ;

Type
  TOF_AFARTICLE_SELECT = Class (TOF_AFTRADUCCHAMPLIBRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

function AFLanceFiche_Mul_Article (strArgument:string):string;

Implementation

uses
{$IFDEF EAGLCLIENT}
        eMul,Maineagl,
{$ELSE}
        Mul, FE_Main,
{$ENDIF}
        UTob, TraducAffaire, dicoaf, UtilMulTrt, afutilarticle;

procedure TOF_AFARTICLE_SELECT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFARTICLE_SELECT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFARTICLE_SELECT.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_AFARTICLE_SELECT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AFARTICLE_SELECT.OnArgument (S : String ) ;
var
   Critere     :string;
//   ExportPath  :string;
//   iPos        :integer;
begin
     inherited;

     // Paramètres: date période
     Critere := Trim (ReadTokenSt (S));
     while Critere <> '' do
     begin
          // Nature de la sélection d'article: pour l'instant, pour la SDA
          if (copy (Critere, 1, 10) = 'NATURE=SDA') then
          begin
               // $$$JP 02/06/03: choix uniquement dans les types prestation, frais et marchandise
               THMultiValComboBox (GetControl('GA_TYPEARTICLE')).Plus  := PlusTypeArticle (TRUE); // ' AND ((CO_CODE="PRE") OR (CO_CODE="MAR") OR (CO_CODE="FRA"))';

               // Traduction du titre de la fiche
               Ecran.Caption := TraduitGA ('Saisie Décentralisée - Sélection des articles');
          end;

          // Paramètre suivant
          Critere := Trim (ReadTokenSt (S));
     end;


{$IFDEF EAGLCLIENT}
     TraduitAFLibGridSt (TFMul(Ecran).FListe);
{$ELSE}
     TraduitAFLibGridDB (TFMul(Ecran).FListe);
{$ENDIF}
end ;

procedure TOF_AFARTICLE_SELECT.OnClose ;
var
   TOBArticle   :TOB;
   F            :TFMul;
begin
     inherited;

     F := TFMul (Ecran);
     if (F.FListe.NbSelected=0) and (not F.FListe.AllSelected) then
        F.Retour := '0'
     else
     begin
          // Récupération des éléments sélectionnés
          TobArticle := TOB.Create ('Les articles', nil, -1);
          TraiteEnregMulListe (F, 'GA_CODEARTICLE', 'ARTICLE', TobArticle, TRUE); //'AFARTICLE_SELECT', TobArticle, True);
          F.Retour := inttostr (integer (TOBArticle));
     end;
end ;

procedure TOF_AFARTICLE_SELECT.OnDisplay () ;
begin
     inherited;
end ;

procedure TOF_AFARTICLE_SELECT.OnCancel () ;
begin
     inherited ;
end ;

function AFLanceFiche_Mul_Article (strArgument:string):string;
begin
     Result := AGLLanceFiche ('AFF', 'AFARTICLE_SELECT', '', '', strArgument);
end;


Initialization
  registerclasses ( [ TOF_AFARTICLE_SELECT ] ) ;
end.

