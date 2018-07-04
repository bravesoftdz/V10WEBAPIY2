{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 26/01/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : GCMULCATALOG_RECH ()
Mots clefs ... : TOF;GCMULCATALOG_RECH
*****************************************************************}
Unit UtofGCMulCatalog_Rech ;

Interface

Uses {StdCtrls, Controls, Classes, db, forms, sysutils, dbTables, ComCtrls,
     HCtrls, HEnt1, HMsgBox, Dialogs, UTOF ; }
     Classes,StdCtrls,HCtrls, UTOF ;

Type
  Tof_GCMulCatalog_Rech = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

var
   MonoArticle:Boolean;

Implementation

procedure TOF_GCMULCATALOG_RECH.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_GCMULCATALOG_RECH.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_GCMULCATALOG_RECH.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_GCMULCATALOG_RECH.OnLoad ;
begin
   Inherited ;
   if MonoArticle then
      BEGIN
      THEdit ( GetControl ( 'GCA_ARTICLE' ) ).Enabled := False ;
      TCheckBox ( GetControl ( '_ARTICLENONREF' ) ).Checked := False ;
      TCheckBox ( GetControl ( '_ARTICLENONREF' ) ).Visible := False ;
      END;
end ;

procedure TOF_GCMULCATALOG_RECH.OnArgument (S : String ) ;
var
  i:Integer;
begin
  Inherited ;
  MonoArticle := False;
  i := Pos(';MULARTICLE',S);
  if (i <> 0) then
     BEGIN
     {Sélection des catalogues d'un article}
     S := Copy(S, 1, i-1);
     MonoArticle := True;
     END;
end ;

procedure TOF_GCMULCATALOG_RECH.OnClose ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_GCMULCATALOG_RECH ] ) ;
end.
