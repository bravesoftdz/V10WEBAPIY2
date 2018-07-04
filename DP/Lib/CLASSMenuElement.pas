{***********UNITE*************************************************
Auteur  ...... : BM
Créé le ...... : 04/09/2002
Modifié le ... : 28/02/2003 : Compatibilité CWAS
Description .. : Classe pour la fusion de documents
Mots clefs ... :
*****************************************************************}

unit CLASSMenuElement;

interface

uses
{$IFDEF EAGLCLIENT}
   eMul,
{$ELSE}
   {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF} Mul,
{$ENDIF}
   HDB, Hctrls, HQry, UTOB, Menus ;


type
   TMenuElement = class

   public

      nMenuDepart_c : integer;
      MyPopUpMenu_c : TPopUpMenu;
      dbgListe_c    : THDBGrid;
      qQueryMul_c   : THQuery;
      frmMul        : TFMul; // MD 01/06/04 : mul appelant pour fetchlestous en eagl

      Procedure InitMenuElement( MyPopUpMenu_p : TPopUpMenu; {dbgListe_p : THDBGrid; qQueryMul_p : THQuery} frm : TFMul ); overload; virtual;
      Procedure InitMenuElement( MyPopUpMenu_p : TPopUpMenu ); overload; virtual;

      procedure InitPopUpElement( MyPopUpMenu_p : TPopUpMenu; bAvecSeparation_p : boolean = false ); overload; virtual;
      procedure FiltreMenu( sRestriction_p : string);
      procedure DetruitMenu( nDepuisIndice_p : integer);
      function  ChoixUnChamp( sChamp_p : string ) : string;
      function  ChargeElementMenu( sRequete_p : string ) : TOB;
      function  OnClickElementMenu(Sender : TObject) : string;

   end;


implementation

uses
     GenMenu, galOutil;

{*****************************************************************
Auteur ....... : BM
Date ......... : 04/09/02
Constructor .. : Create
Description .. : Construit et initialise la classe
Paramètres ... : Le menu POPUP
                 le bouton BCHERCHE
                 la fiche (Mul)
                 // la grille de la fiche
                 // Le query Q de la fiche
*****************************************************************}

Procedure TMenuElement.InitMenuElement( MyPopUpMenu_p : TPopUpMenu; {dbgListe_p : THDBGrid;
               qQueryMul_p : THQuery;} frm : TFMul);
begin
   MyPopUpMenu_c := MyPopUpMenu_p;
   frmMul := frm;
   dbgListe_c := frm.FListe; // dbgListe_p;
   qQueryMul_c := frm.Q; // qQueryMul_p;
   InitPopUpElement( MyPopUpMenu_p );
end;

Procedure TMenuElement.InitMenuElement( MyPopUpMenu_p : TPopUpMenu );
begin
   MyPopUpMenu_c := MyPopUpMenu_p;
   nMenuDepart_c := 0;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 18/09/02
Procédure .... : DetruitPopUp
Description .. : A la fermeture de la fenêtre
Paramètres ... :
*****************************************************************}

procedure TMenuElement.DetruitMenu( nDepuisIndice_p : integer);
begin
   DeletePopUp( MyPopUpMenu_c, nMenuDepart_c + nDepuisIndice_p);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 20/11/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TMenuElement.FiltreMenu( sRestriction_p : string);
begin
   FiltrePopUp( MyPopUpMenu_c, nMenuDepart_c, sRestriction_p );
end;
{*****************************************************************
Auteur ....... : BM
Date ......... : 18/09/02
Procédure .... : InitPopUpElement
Description .. : Initialise le popup menu
Paramètres ... :
*****************************************************************}

procedure TMenuElement.InitPopUpElement( MyPopUpMenu_p : TPopUpMenu; bAvecSeparation_p : boolean = false );
begin
   if bAvecSeparation_p then
      nMenuDepart_c := InitPopUpSubMenu( MyPopUpMenu_p, '-' ) - 1
   else
      nMenuDepart_c := MyPopUpMenu_p.Items.Count - 1;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 18/09/02
Procédure .... : ChoixUnChamp
Description .. : Le champ à traiter
Paramètres ... :
*****************************************************************}

function TMenuElement.ChoixUnChamp( sChamp_p : string ) : string;
var
   sTitre_l : string;
   nIDField_l, nPos_l : integer;
begin
   sTitre_l := '';
   //CWASERREUR en attendant galoutil...
   {$IFDEF EAGLCLIENT}
      sTitre_l := ChampToLibelle(sChamp_p);
      nPos_l := Pos ( ':', sTitre_l );
      if nPos_l > 0 then
         sTitre_l := Copy ( sTitre_l, 1, nPos_l - 1 ) ;
  {$ELSE}
   if not ChampEstDansQuery( sChamp_p, qQueryMul_c) then
   begin
      sTitre_l := ChampToLibelle(sChamp_p);
      nPos_l := Pos ( ':', sTitre_l );
      if nPos_l > 0 then
         sTitre_l := Copy ( sTitre_l, 1, nPos_l - 1 ) ;
   end
   else
   begin
      nIDField_l := qQueryMul_c.FieldByName(sChamp_p).Index - 1;
      sTitre_l := dbgListe_c.Columns.Items[nIDField_l].Title.Caption;
   end;
   {$ENDIF}
   result := sTitre_l;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 18/09/02
Procedure .... : ChargeTypeTaches
Description .. : Charge les types d'évènements
Paramètres ... : TOB évènement
                 Le menu popup
*****************************************************************}

function  TMenuElement.ChargeElementMenu( sRequete_p : string ) : TOB;
var
   TOBElement_l : TOB;
   qQuery_l     : TQuery;

begin
   TOBElement_l := TOB.Create('Element', NIL, -1);
   qQuery_l := OpenSQL( sRequete_p, true );
   TOBElement_l.LoadDetailDB ('Element', '', '', qQuery_l, FALSE, FALSE) ;
   Ferme (qQuery_l);
   result := TOBElement_l;
end;


{*****************************************************************
Auteur ....... : BM
Date ......... : 18/09/02
Procédure .... : OnClickElement
Description .. : Procédure exécutée lors du click sur les éléments
                 du sous menu
Paramètres ... : L'objet
*****************************************************************}

function TMenuElement.OnClickElementMenu(Sender : TObject) : string;
//var
//   nSousMenuIndex_l : integer;
begin
//   nSousMenuIndex_l := (Sender as TMenuItem).MenuIndex;
   result := ( Sender as TMenuItem ).Hint;
end;



end.
