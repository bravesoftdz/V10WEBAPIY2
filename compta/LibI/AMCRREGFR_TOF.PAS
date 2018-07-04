{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 22/10/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AMCRREGFR ()
Mots clefs ... : TOF;AMCRREGFR
*****************************************************************}
Unit AMCRREGFR_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
     PrintDbg,
{$else}
     eMul,
     MaineAGL,
     UtileAGL,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     Vierge,
     UTOF ;

Type
  TOF_AMCRREGFR = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
     { Variables }
     fGrille: THGrid;
     fValid : string;
     { Procedures }
     procedure OnClickBImprimer(Sender: TObject);
  end ;

Function AMLanceCompteRendu_REGFR(stParam: string) : integer;


Implementation

Function AMLanceCompteRendu_REGFR(stParam: string) : integer;
begin
  result:= StrToInt(AGLLanceFiche('AM', 'AMCRREGFR','' ,'',stParam) );
end;

procedure TOF_AMCRREGFR.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AMCRREGFR.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AMCRREGFR.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_AMCRREGFR.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AMCRREGFR.OnArgument (S : String ) ;
begin
  Inherited ;
  TFVierge(Ecran).retour := IntToStr(MrYes);
  TToolbarButton97(GetControl('BImprimer')).Onclick := OnClickBimprimer;

  // Récup paramètre en entrée
  fValid := ReadTokenSt(S);
  SetControlText ('NBValid', 'Nombre de fiches validées : ' + fValid);

  // Récup zones
  SetControlVisible ('BValider', False);
  fGrille := THGrid (GetControl ('CPTEListe'));
  fGrille.ColAligns[0]:= taCenter ;

  // Remplissage THGRID par LaTob <=> TheTob globale qui pointe depuis l'appelant
  LaTob.PutGridDetail(fGrille, False, False, 'COMPTE');
end ;


procedure TOF_AMCRREGFR.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_AMCRREGFR.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AMCRREGFR.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_AMCRREGFR.OnClickBImprimer(Sender: TObject);
Var
  Titre : String ;
begin
  Titre := TraduireMemoire('Compte-rendu de validation');

  // Lignes supplémentaires pour présenter le nb de fiches traitées
  fGrille.CellValues[0,fGrille.RowCount]:= '#FONT#|clRed|G#Nb fiches validées :';
  FGrille.CellValues[0, fGrille.RowCount+1]:= fValid;
  fGrille.RowCount := fGrille.RowCount + 2;

 {$IFDEF EAGLCLIENT}
     PrintDBGrid(Titre, FGrille.ListeParam, '', '');
 {$ELSE}
     PrintDBGrid(FGrille, nil, Titre, '') ;
 {$ENDIF}

  // Retirer les lignes supplémentaires
  fGrille.RowCount := fGrille.RowCount - 2;
end;

Initialization
  registerclasses ( [ TOF_AMCRREGFR ] ) ;
end.
