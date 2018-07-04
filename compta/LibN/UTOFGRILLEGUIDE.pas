{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : CPTRESOGUIDE ()
Mots clefs ... : TOF;CPTRESOGUIDE
*****************************************************************}
Unit UTOFGRILLEGUIDE ;

Interface

Uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
     HCtrls, HEnt1, HMsgBox,
     // VCL
     Windows,
     Grids,
     // composant AGL
     HSysMenu,
     HTB97,
     // lib
     AGLInit,
     {$IFNDEF EAGLCLIENT}
     FE_Main,   // pour AGLLanceFiche()
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     {$ELSE}
     MaineAGL,
     {$ENDIF}
     UTOB,
{$IFDEF VER150}
     variants,
{$ENDIF}
     UTOF ;

Type

  TOF_GRILLEGUIDE = Class (TOF)
    private
     // variable interne
     FTOB                    : TOB;
     E                       : TControl;
     HSystemMenu1            : THSystemMenu;
     // liste des contrôles
     HGSaisie                : THGrid;                 // Grille de saisie
     FChamps1                : string;
     FChamps2                : string;
     BValider                : TToolbarButton97;       // Bouton Valider : coche verte

     procedure BValiderClick(Sender: TObject);
     procedure PositionneLookUp;         // Bouton Valider

    public
     procedure OnArgument (S : String ) ; override ;
     procedure OnClose                  ; override ;

  end ;

function LookUpTob ( E : TControl ; vTOB : TOB ; Titre, Colonne, Entete : string ) : TOB;

Implementation

Const

 cColCode = 0;
 cColLib  = 1;
 cColCrit = 2;

function LookUpTob ( E : TControl ; vTOB : TOB ; Titre, Colonne,Entete : string ) : TOB;
var
 lData    : TObject;
 lTOB     : TOB;
begin

 // sauvegarde des valeurs initiales
 lTOB                := TheTOB;
 lData               := TheData;
 // affectation avec les valeurs envoyée
 TheTOB              := vTOB;
 TheData             := E;

 V_PGI.FormCenter    := false;
 AGLLanceFiche('CP','RLBGUIDE','','',Titre + ';' + Colonne + ';' + Entete) ;
 V_PGI.FormCenter    := true;

 // on recupere la TheTOB qui contient la lignes selectionné
 Result              := TheTOB;
 // reaffectation des valeurs initiales
 TheTOB              := lTOB;
 TheData             := lData;

end;

{ ZGRILLEGUIDE }

procedure TOF_GRILLEGUIDE.OnArgument( S : String );
var
 lStArg   : string;
 i        : integer;
 lTOB     : TOB;
begin

 inherited;
 // creation des controles
 FTOB                                 := TOB(laTOB);
 E                                    := TControl(TheData);
 HSystemMenu1                         := THSystemMenu.Create(Ecran);

 HGSaisie                             := THGrid(GetControl('FE__HGRID_SAISIE'));
 BValider                             := TToolbarButton97(GetControl('BValider')) ;

 lStArg                               := S;

 Ecran.Caption                        := TraduireMemoire(ReadTokenST(lStArg));
 UpdateCaption(Ecran);

 FChamps1                             := ReadTokenST(lStArg);
 FChamps2                             := ReadTokenST(lStArg);

 HGSaisie.Cells[cColCode,0]           := TraduireMemoire(ReadTokenST(lStArg));
 HGSaisie.Cells[cColLib,0]            := TraduireMemoire(ReadTokenST(lStArg));

 HGSaisie.RowCount                    := FTOB.Detail.Count + 1;

 // remplissage de la grille
 for i := 0 to FTOB.Detail.Count - 1 do
  begin

   lTOB := FTOB.Detail[i];
   HGSaisie.Row                      := i + 1;
   HGSaisie.Cells[cColCode,i+1]      := varToStr(lTOB.GetValue(FChamps1));
   HGSaisie.Cells[cColLib,i+1]       := varToStr(lTOB.GetValue(FChamps2));

   HGSaisie.Objects[0,HGSaisie.Row]  := lTOB;

  end; // for

 // redimensionnement de la grille
  HSystemMenu1.ResizeGridColumns(HGSaisie);

 // Affiche la fenêtre en dessous du contrôle E
 PositionneLookUp;

 // on se positionne sur la premiere ligne
 HGSaisie.Row := 1;

 // assignation des eve
 BValider.OnClick                    := BValiderClick;

end;

procedure TOF_GRILLEGUIDE.OnClose;
begin

 if assigned(HSystemMenu1)  then HSystemMenu1.Free;

 HSystemMenu1 := nil;

 inherited;

end;


procedure TOF_GRILLEGUIDE.BValiderClick( Sender : TObject );
begin
 TheTOB := TOB( HGSaisie.Objects[0,HGSaisie.Row] );
end;


procedure TOF_GRILLEGUIDE.PositionneLookUp;
var
 R  : TRect ;
 PP : TPoint ;
begin

 if E is TStringGrid then
  begin
   R           := TStringGrid(E).CellRect(TStringGrid(E).Col,TStringGrid(E).Row);
   R.Left      := R.Left     + TStringGrid(E).Left;
   R.Top       := R.Top      + TStringGrid(E).Top;
   R.Bottom    := R.Bottom   + TStringGrid(E).Top;
   R.Right     := R.Right    + TStringGrid(E).Left;
  end
   else
     R := E.BoundsRect;


 PP.Y := R.Bottom;
 PP.X := R.Left;
 PP   := E.parent.ClientToScreen(PP);

 if (PP.X + Ecran.Width) > Screen.Width then
  PP.X := PP.X - Ecran.Width + R.Right -R.Left;

 if ( PP.Y + Ecran.Height ) > Screen.Height then
  begin
   if PP.Y - Ecran.Height - (R.Bottom-R.Top) < 0 then
    PP.Y := ( Screen.Height - Ecran.Height ) div 2
     else
      PP.Y := PP.Y - Ecran.Height - ( R.Bottom - R.Top);
   end ;

 Ecran.Left := PP.X;
 Ecran.Top  := PP.Y;

end ;


Initialization
  registerclasses ( [ TOF_GRILLEGUIDE ] ) ;
end.

