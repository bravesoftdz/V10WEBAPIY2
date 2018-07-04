{***********UNITE*************************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... : 09/04/2003
Description .. : Affichage d'une fenêtre avec la taille du fichier .TOX 
Suite ........ : lorsqu'on clique sur le fichier avec le bouton droit de la 
Suite ........ : souris et que celui-ci est sélectionné
Mots clefs ... : 
*****************************************************************}
unit PaieTOXInfoTox;

interface

uses
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  StdCtrls,
  HPanel,
  UIUtil ;

type
  TFicheInfoTox = class(TForm)
    LLibelle: TLabel;
    LTaille: TLabel;
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

procedure InfoTox ( sFile : string ; Taille : longint ) ;

var
  FicheInfoTox: TFicheInfoTox;

implementation

{$R *.DFM}

procedure InfoTox ( sFile : string ; Taille : longint ) ;
var X  : TFicheInfoTox ;
    PP : THPanel ;
begin
PP:=FindInsidePanel ;
X:=TFicheInfoTox.Create(Application) ;
if PP=Nil then
  begin
  Try
    X.Caption      := X.Caption + ' ' + sFile ;
    X.LTaille.Caption := FormatFloat ( '#,###,###,##0', Taille ) ;
    X.ShowModal;
  finally
    X.Free;
  end;
   end else
   begin
     InitInside(X,PP) ;
     X.Show ;
   end ;
end ;

end.
