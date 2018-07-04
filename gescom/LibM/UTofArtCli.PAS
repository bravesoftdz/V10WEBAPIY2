{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 18/06/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : ARTCLI ()
Mots clefs ... : TOF;ARTCLI
*****************************************************************}
Unit UTofArtCli ;

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
     UTOF,
     AGLInitGC,
     UTOB,UTOM,
HDimension,
      Dialogs,M3FP,EntGC;

Type
  TOF_ARTCLI = Class (TOF)
    procedure OnLoad ; override;
  end ;
Implementation

procedure TOF_ARTCLI.OnLoad;
var CodeArticle : THEdit;
begin
  inherited;
  CodeArticle:=THEdit(GetControl('GL_CODEARTICLE')) ;
  CodeArticle.Plus := ' AND GL_TYPEARTICLE="MAR"' ;
  CodeArticle.Plus := CodeArticle.Plus + ' AND GL_NATUREPIECEG="'+THValComboBox(GetControl('GL_NATUREPIECEG')).Value;
  CodeArticle.Plus := CodeArticle.Plus + '" AND GL_TIERS = "'+THEdit(GetControl('GL_TIERS')).Text+'"' ;
end;


{ TOF_ARTCLI }

Initialization
  registerclasses ( [ TOF_ARTCLI ] ) ;
end.
