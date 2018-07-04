{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : CPTRESOGUIDE ()
Mots clefs ... : TOF;CPTRESOGUIDE
*****************************************************************}
Unit UTOFLOOKUPTOB ;

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
     Hdb,
     {$ELSE}
     MaineAGL,
     {$ENDIF}
     UTOB,
     {$IFDEF VER150}
     Variants,
     {$ENDIF}
     Vierge,
     Dialogs, // pour TFinddialog
     UTOF ;

Type

  TOF_LOOKUPTOB = Class (TOF)
    private
     // variable interne
     FInNbChamps             : integer ;
     FTOB                    : TOB;
     E                       : TControl;
     HSystemMenu1            : THSystemMenu;
     // liste des contrôles
     GS                      : THGrid;                 // Grille de saisie
     FChamps1                : string;
     FChamps2                : string;
     FChamps3                : string;
     FBoRemplirControl       : boolean;
     FStTag                  : string;
     BValider                : TToolbarButton97;       // Bouton Valider : coche verte
     FindSais                : TFindDialog;

     procedure BValiderClick(Sender: TObject);
     procedure PositionneLookUp;         // Bouton Valider
     procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
     procedure GSDblClick( Sender : TObject) ;
     procedure AfterShow ;
     procedure SearchClick;
     procedure FindSaisFind (Sender : TObject);

    public
     procedure OnArgument (S : String ) ; override ;
     procedure OnClose                  ; override ;

  end ;

function LookUpTob ( E : TControl ; vTOB : TOB ; Titre, Colonne, Entete : string ; RemplirControl : boolean = false ; Tag : string = '' ) : TOB;

Implementation

Const

 cColCode = 0;
 cColLib  = 1;
 cColCrit = 2;

function LookUpTob ( E : TControl ; vTOB : TOB ; Titre, Colonne,Entete : string ; RemplirControl : boolean = false ; Tag : string = '' ) : TOB;
var
 lData    : TObject;
 lTOB     : TOB;
 lStControl : string ;
begin

 // sauvegarde des valeurs initiales
 lTOB                := TheTOB;
 lData               := TheData;
 // affectation avec les valeurs envoyée
 TheTOB              := vTOB;
 TheData             := E;
 if RemplirControl then lStControl := 'X' else lStControl := '-' ;

 V_PGI.FormCenter    := false;
 AGLLanceFiche('CP','CLOOKUPTOB','','',Titre + '|' + Colonne + '|' + Entete + '|' + lStControl + '|' + Tag) ;
 V_PGI.FormCenter    := true;

 // on recupere la TheTOB qui contient la lignes selectionné
 Result              := TheTOB;
 // reaffectation des valeurs initiales
 TheTOB              := lTOB;
 TheData             := lData;

end;

{ TOF_LOOKUPTOB }

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 21/10/2003
Modifié le ... :   /  /
Description .. : - 22/10/2003 - LG - FB 12905 - correction pour la gestion
Suite ........ : des valurs afficher, ajout d'un filtre sur les enregsitrement a
Suite ........ : afficher
Mots clefs ... :
*****************************************************************}
procedure TOF_LOOKUPTOB.OnArgument( S : String );
var
 lStArg    : string;
 i         : integer;
 lTOB      : TOB;
 lIndex    : integer ;
 lStTitre  : string ;
 lStCol    : string ;
 lStLib    : string ;
 lIndexVal : integer ;
 lValeur   : string ;
begin

 inherited;
 // creation des controles
 FTOB            := TOB(laTOB);

 E               := TControl(TheData);
 FindSais        := TFindDialog.Create(Ecran) ;
 FindSais.OnFind := FindSaisFind ;
 lIndexVal       := 1 ;

 if ( FTOB = nil ) then
  begin
   MessageAlerte('TOF_LOOKUPTOB.OnArgument : la TOB = nil ' ) ;
   exit ;
  end;

 if ( FTOB.Detail = nil ) or ( FTOB.Detail.Count = 0 ) then
  begin
   MessageAlerte('Pas de données a afficher');
   // Affiche la fenêtre en dessous du contrôle E
   PositionneLookUp;
   Ecran.Close ;
   exit ;
  end;

  if E <> nil then
   begin // on recupere la valeur du controle
    if E is TStringGrid then
     lValeur := UpperCase(THGrid(E).Cells[THGrid(E).Col,THGrid(E).Row] )
      else
       if ( E is THEdit ) then
        lValeur := UpperCase(THEdit(E).Text)
         else
          if ( E is THCritMaskEdit ) then
           lValeur := UpperCase(THCritMaskEdit(E).Text)
            else
             lValeur := '' ;
   end ;


 FInNbChamps                          := 0 ;

 lStArg                               := S;
 HSystemMenu1                         := THSystemMenu.Create(Ecran);
 GS                                   := THGrid(GetControl('G'));
 BValider                             := TToolbarButton97(GetControl('BValider'))  ;
 GS.OnDblClick                        := GSDblClick ;
 Ecran.OnKeyDown                      := FormKeyDown;

 lIndex                               := Pos('|',lStArg) ;
 lstTitre                             := Copy(lStArg,1,lIndex-1) ;
 lStArg                               := Copy(lStArg,lIndex+1,length(lStArg)) ;
 lIndex                               := Pos('|',lStArg) ;
 lStCol                               := Copy(lStArg,1,lIndex-1) ;
 lStArg                               := Copy(lStArg,lIndex+1,length(lStArg)) ;
 lIndex                               := Pos('|',lStArg) ;
 lStLib                               := Copy(lStArg,1,lIndex-1) ;
 lStArg                               := Copy(lStArg,lIndex+1,length(lStArg)) ;
 lIndex                               := Pos('|',lStArg) ;
 FBoRemplirControl                    := Copy(lStArg,1,1) = 'X' ;
 FStTag                               := Copy(lStArg,lIndex+1,1) ;

 Ecran.Caption                        := lstTitre ;
 UpdateCaption(Ecran);

 FChamps1                             := ReadTokenST(lStCol) ;
 FChamps2                             := ReadTokenST(lStCol) ;
 FChamps3                             := ReadTokenST(lStCol) ;

 GS.Cells[cColCode,0]                 := TraduireMemoire(ReadTokenST(lStLib));
 Inc(FInNbChamps) ;

 if lStLib <> '' then
  begin
   GS.Cells[cColLib,0]                := TraduireMemoire(ReadTokenST(lStLib));
   Inc(FInNbChamps) ;
  end; // if

 if lStLib <> '' then
  begin
   GS.Cells[cColLib+1,0]                := TraduireMemoire(ReadTokenST(lStLib));
   Inc(FInNbChamps) ;
  end; // if

 GS.ColCount                          := FInNbChamps ;
 GS.RowCount                          := 2 ;
 lIndex                               := 0 ;


 // remplissage de la grille
 for i := 0 to FTOB.Detail.Count - 1 do
  begin

   lTOB := FTOB.Detail[i];
   if ( FStTag <> '' ) and ( FTOB.Detail[i].GetValue('TAG') <> FStTag ) then continue ;
   Inc(lIndex) ;
   if ( lIndex ) >= GS.RowCount then
    GS.RowCount := GS.RowCount + 1 ;
   GS.Row                            := lIndex ;
   GS.Cells[cColCode,lIndex]         := varToStr(lTOB.GetValue(FChamps1));
   if FInNbChamps >= 2 then
    GS.Cells[cColLib,lIndex]         := varToStr(lTOB.GetValue(FChamps2));
   if FInNbChamps = 3 then
    GS.Cells[cColLib+1,lIndex]        := varToStr(lTOB.GetValue(FChamps3));

   if (lValeur = UpperCase(GS.Cells[cColLib,lIndex]) ) or (lValeur = GS.Cells[cColCode,lIndex]) then
     lIndexVal := lIndex ;

   GS.Objects[0,GS.Row]              := lTOB ;

  end; // for

 // Affiche la fenêtre en dessous du contrôle E
// PositionneLookUp;

 TFVierge(Ecran).OnAfterFormShow := AfterShow ;

 // on se positionne sur la premier ligne ou l'enregistrement que l'on a retrouver
 //GS.Row                              := 1;
 GS.Row := lIndexVal ;

 // assignation des eve
 BValider.OnClick                   := BValiderClick;
  if GS.ColCount = 3 then  // Fiche 18238
  begin
      Ecran.Width := 400;
      TFVierge(Ecran).HMTrad.ResizeGridColumns(GS) ;
  end;
// TFVierge(Ecran).HMTrad.ResizeGridColumns(GS) ;

end;

procedure TOF_LOOKUPTOB.OnClose;
begin

 if assigned(HSystemMenu1)  then HSystemMenu1.Free;
 if assigned(FindSais)      then FindSais.Free;

 FindSais := nil;

 inherited;

end;

procedure TOF_LOOKUPTOB.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin

  if ( csDestroying in Ecran.ComponentState ) then Exit ;

  case Key of
   VK_F10  : begin
                  BValiderClick(nil);
                  Ecran.Close;
             end; // VK_F10
{CTRL+F} 70 : if Shift=[ssCtrl] then begin Key:=0 ; SearchClick ; end ;
  end; // if

end;

procedure TOF_LOOKUPTOB.GSDblClick( Sender : TObject) ;
begin
 BValiderClick(nil) ;
 Ecran.Close ;
end;


procedure TOF_LOOKUPTOB.BValiderClick( Sender : TObject );
begin
 TheTOB := TOB( GS.Objects[0,GS.Row] );

 if E = nil then exit ;

 if FBoRemplirControl and ( E is THGrid ) then
  THGrid(E).Cells[THGrid(E).Col,THGrid(E).Row] := GS.Cells[GS.Col,GS.Row]
  else
    if FBoRemplirControl and ( E is THEdit ) then
      THEdit(E).Text := GS.Cells[GS.Col,GS.Row]
       else
        if FBoRemplirControl and ( E is THCritMaskEdit ) then
         THCritMaskEdit(E).Text := GS.Cells[GS.Col,GS.Row]

{$IFDEF EAGLCLIENT}
    ;
{$ELSE}
    else
      if FBoRemplirControl and ( E is THDBEdit ) then
        THDBEdit(E).Text := GS.Cells[GS.Col,GS.Row];
{$ENDIF}
end;


procedure TOF_LOOKUPTOB.PositionneLookUp;
var
 R  : TRect ;
 PP : TPoint ;
begin

 if E = nil then exit ;

 if E is TStringGrid then
  begin
   R           := TStringGrid(E).CellRect(TStringGrid(E).Col,TStringGrid(E).Row);
   R.Left      := R.Left     + TStringGrid(E).Left;
   R.Top       := R.Top      + TStringGrid(E).Top;
   R.Bottom    := R.Bottom   + TStringGrid(E).Top;
   R.Right     := R.Right    + TStringGrid(E).Left;
  end
   else
    begin
     R := E.BoundsRect;
    end ;


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


procedure TOF_LOOKUPTOB.AfterShow;
begin
 PositionneLookUp ;
end;

procedure TOF_LOOKUPTOB.SearchClick ;
begin
 FindSais.Execute ;
end ;

procedure TOF_LOOKUPTOB.FindSaisFind(Sender: TObject) ;
var
 FindFirst : boolean ;
begin
 Rechercher(GS, FindSais, FindFirst ) ;
end ;

Initialization
  registerclasses ( [ TOF_LOOKUPTOB ] ) ;
end.

