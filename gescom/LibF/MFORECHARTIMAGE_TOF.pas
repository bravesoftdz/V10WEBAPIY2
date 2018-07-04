{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 28/10/2003
Modifié le ... : 28/10/2003
Description .. : Source TOF de la FICHE : MFORECHARTIMAGE
Suite ........ :
Suite ........ : Recherche des articles par image
Mots clefs ... : TOF;MFORECHARTIMAGE
*****************************************************************}
Unit MFORECHARTIMAGE_TOF ;

Interface

uses
  Classes, forms, sysutils, Controls,
{$IFDEF EAGLCLIENT}
  Maineagl,
{$ELSE}
  Fe_Main, UTofMulImage,
{$ENDIF}
  HEnt1, HCtrls, UTOF;

Type
{$IFDEF EAGLCLIENT}
  TOF_MFORECHARTIMAGE = Class (TOF)
{$ELSE}
  TOF_MFORECHARTIMAGE = Class (TOF_MULIMAGE)
{$ENDIF}
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

function FOLanceFicheRechArtImage(Range, Lequel, Argument: string): string;

Implementation

uses
  UtilGC;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 28/10/2003
Modifié le ... : 28/10/2003
Description .. : Lancement de la fiche de la recherche des articles par image
Mots clefs ... : 
*****************************************************************}

function FOLanceFicheRechArtImage(Range, Lequel, Argument: string): string;
begin
  Result := AGLLanceFiche('MFO', 'MFORECHARTIMAGE', Range, Lequel, Argument);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 28/10/2003
Modifié le ... : 28/10/2003
Description .. : OnNew
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORECHARTIMAGE.OnNew ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 28/10/2003
Modifié le ... : 28/10/2003
Description .. : OnDelete
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORECHARTIMAGE.OnDelete ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 28/10/2003
Modifié le ... : 28/10/2003
Description .. : OnUpdate
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORECHARTIMAGE.OnUpdate ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 28/10/2003
Modifié le ... : 28/10/2003
Description .. : OnLoad
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORECHARTIMAGE.OnLoad ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 28/10/2003
Modifié le ... : 28/10/2003
Description .. : OnArgument
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORECHARTIMAGE.OnArgument (S : String ) ;
var
  Ind: integer;
  Stg: string;
begin
  Inherited ;
  MajChampsLibresArticle(TForm(Ecran));
  for Ind := 1 to 3 do
  begin
    Stg := RechDom('GCLIBFAMILLE', 'LF'+ InttoStr(Ind), False);
    SetControlText('TGA_FAMILLENIV'+ InttoStr(Ind), Stg);
  end;
  if V_PGI.LaSerie = S3 then
  begin
    SetControlVisible('GA_LOT', False);
    SetControlVisible('PZONESLIBRES', False);
    for Ind := 4 to 9 do
    begin
      SetControlVisible('GA_LIBREART'+ InttoStr(Ind), False);
      SetControlVisible('TGA_LIBREART'+ InttoStr(Ind), False);
    end;
    SetControlVisible('GA_LIBREARTA', False);
    SetControlVisible('TGA_LIBREARTA', False);
  end;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 28/10/2003
Modifié le ... : 28/10/2003
Description .. : OnClose
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORECHARTIMAGE.OnClose ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 28/10/2003
Modifié le ... : 28/10/2003
Description .. : OnDisplay
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORECHARTIMAGE.OnDisplay () ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 28/10/2003
Modifié le ... : 28/10/2003
Description .. : OnCancel
Mots clefs ... :
*****************************************************************}

procedure TOF_MFORECHARTIMAGE.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_MFORECHARTIMAGE ] ) ;
end.
