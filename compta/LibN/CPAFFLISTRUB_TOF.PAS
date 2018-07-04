{***********UNITE*************************************************
Auteur  ...... : Jean Pastéris
Créé le ...... : 20/06/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPAFFLISTRUB ()
Mots clefs ... : TOF;CPAFFLISTRUB
*****************************************************************}
Unit CPAFFLISTRUB_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     MaineAGL,
     UtileAgl,
{$ELSE}
     FE_Main,
     PrintDBG,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HTB97,
     UTOB,
     HEnt1,
     HMsgBox,
     UTOF ;

Type
  TOF_CPAFFLISTRUB = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
  private
    Titre  : string;
    FListe : THGrid;
    procedure BImprimerClick(Sender: TObject);
  end ;

procedure AfficheListe(LaListe: Tob; LeTitre: String; const LesTitresGrid: array of string);


Implementation

var
  MaTob : TOB;
  TitresListe : array[0..2] of string;

{---------------------------------------------------------------------------------------}
procedure AfficheListe(LaListe: Tob; LeTitre: String; const LesTitresGrid: array of string);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  if LaListe = nil then Exit ;

  MaTob := LaListe;


  for n := 0 to High(LesTitresGrid) do
    TitresListe[n] := LesTitresGrid[n];
  TitresListe[2] := 'Description';

  AGLLanceFiche('CP', 'CPAFFLISTRUB', '', '', 'ACTION=CONSULTATION' + ';' + LeTitre +';');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPAFFLISTRUB.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  Inherited ;
  TToolbarButton97(GetControl('BIMPRIMER')).OnClick := BImprimerClick;
  FListe := THGrid(GetControl('FLISTE'));
  ReadTokenSt(S);
  Titre := ReadTokenSt(S);
  SetControlEnabled('BVALIDER', True);

  for n := 0 to High(TitresListe) do
    Fliste.Cells[n, 0] := TitresListe[n] ;

  MaTOb.PutGridDetail(Fliste,False,True,'',true) ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPAFFLISTRUB.BImprimerClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
{$IFDEF EAGLCLIENT}
  LanceEtatTOB('E', 'RAS', 'RAS', MaTob, True, False, False, nil, '', 'Liste des rubriques créées', False);
{$ELSE}
  PrintDBGrid(FListe, nil, Ecran.Caption, '') ;
{$ENDIF}
end ;


Initialization
  registerclasses ( [ TOF_CPAFFLISTRUB ] ) ;

end.
