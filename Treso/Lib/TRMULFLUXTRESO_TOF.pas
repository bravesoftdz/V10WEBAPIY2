{ Unité : Source TOF de la FICHE : TRMULFLUXTRESO
--------------------------------------------------------------------------------------
    Version  |   Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 0.91         06/11/03   JP   Création de l'unité
1.01.001.001  04/03/04   JP   Mise en place d'un contrôle dans le OnDeleteRecor sur
                              le Code Flux pour interdire les modifications sur les 3
                              flux EQD, EQR, REI fournis dans la "maquette"
6.30.001.004  15/03/05   JP   Modification de SupprimerClik pour l'adapter à FQ 10221
7.05.001.001  12/10/06   JP   Filtre sur les classes de Flux (Multi sociétés)
-------------------------------------------------------------------------------------}
unit TRMULFLUXTRESO_TOF ;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  Mul, FE_Main,
  {$ELSE}
  eMul, MaineAGL,
  {$ENDIF}
  Forms, HCtrls, HMsgBox, UTOF, HTB97;


type
  TOF_TRMULFLUXTRESO = class (TOF)
    procedure OnArgument(S : string); override ;
  private
    procedure FListeDblClick(Sender : TObject);
    procedure NouveauClick  (Sender : TObject);
    procedure SupprimerClik (Sender : TObject);
  end ;

procedure TRLanceFiche_MulFluxTreso(Dom, Fiche, Range, Lequel, Arguments : string);

implementation

uses TomFluxTreso, Constantes, Commun;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_MulFluxTreso(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULFLUXTRESO.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 150;
  TFMul(Ecran).FListe.OnDblClick := FListeDblClick;
  TFMul(Ecran).FListe.TwoColors := True;
  {$IFDEF EAGLCLIENT}
  TFMul(Ecran).FListe.MultiSelect := True;
  {$ELSE}
  TFMul(Ecran).FListe.MultiSelection := True;
  {$ENDIF}
  TFMul(Ecran).Binsert.OnClick := NouveauClick;
  TToolbarButton97(GetControl('BDELETE')).OnClick := SupprimerClik;
  {12/10/06 : On exclut la Classe ICC si on n'est pas en MultiDossier}
  if not IsTresoMultiSoc then
    THValComboBox(GetControl('TFT_CLASSEFLUX')).Plus := 'AND CO_CODE <> "' + cla_ICC + '"'; 
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULFLUXTRESO.FListeDblClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TRLanceFiche_FluxTreso('TR','TRFLUXTRESO', '', GetField('TFT_FLUX'), 'ACTION=MODIFICATION');
  TFMul(Ecran).BCherche.Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULFLUXTRESO.NouveauClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TRLanceFiche_FluxTreso('TR','TRFLUXTRESO', '', '', 'ACTION=CREATION');
  TFMul(Ecran).BCherche.Click;
end;

{15/03/05 : Un seul message pour toute la boucle de suppression
            Utilisation de ExisteSQL plutôt quer OpenSQL
            Adaptation de la fonction aux modifications de la FQ 10221
{---------------------------------------------------------------------------------------}
procedure TOF_TRMULFLUXTRESO.SupprimerClik(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  f : string;
  n : Integer;
  Err : Boolean;
begin
  {JP 11/03/04 : Il n'y avait même pas de message de confirmation !! Mise en conformité
                 avec les règles d'ergonomie PGI}
  if TFMul(Ecran).FListe.nbSelected = 0 then begin
    HShowMessage('0;' + Ecran.Caption + ';Veuillez sélectionner un flux.;W;O;O;O;', '', '');
    Exit;
  end;

  {03/06/04 : demande de confirmation de suppression}
  if HShowMessage('1;' + Ecran.Caption + ';Êtes-vous sûr de vouloir supprimer les flux sélectionnés ?;Q;YN;N;N;', '', '') = mrNo then
    Exit;

  {15/03/05 : Un seul message pour toute la boucle}
  Err := False;

  for n := 0 to TFMul(Ecran).FListe.nbSelected - 1 do begin
    TFMul(Ecran).FListe.GotoLeBookmark(n);
    {JP 04/03/04 : On interdit la modification des enregistrements fournis dans la maquette}
    if (GetField('TFT_CLASSEFLUX') = cla_Reference) then begin
      Err := True;
      Continue;
    end;

    f := VarToStr(GetField('TFT_FLUX'));

    {On s'assure que les flux n'est pas utilisé dans les frais}
    if ExisteSql('SELECT TFR_CODEFLUX FROM FRAIS WHERE TFR_CODEFLUX = "' + f + '"') then begin
      Err := True;
      Continue;
    end;

    {On s'assure que le flux n'est pas utilisé dans les catégories}
    if ExisteSql('SELECT TTR_TRANSAC FROM TRANSAC WHERE TTR_VERSEMENT = "' + f + '" OR TTR_AMORTISSEMENT = "' + f +
                 '" OR TTR_AGIOSINTERETS = "' + f + '" OR TTR_REMBANTICIPE = "' + f + '" OR TTR_PENREMBANTIC = "' + f + '"') then begin
      Err := True;
      Continue;
    end;

    {On s'assure que le flux n'est pas utilisé dans les écritures}
    if ExisteSql('SELECT TE_CODEFLUX FROM TRECRITURE WHERE TE_CODEFLUX = "' + f + '"') then begin
      Err := True;
      Continue;
    end;

    ExecuteSQL('DELETE FROM FLUXTRESO WHERE TFT_FLUX = "' + f + '"');
  end;

  {15/03/05 : Un seul message pour toute la boucle}
  if Err then
    HShowMessage('0;' + Ecran.Caption + ';Certains flux n''ont pu être supprimés :'#13 +
                 ' - Soit parce qu''ils sont de classe "référence"'#13 +
                 ' - Soit parce qu''ils sont utilisés par des écritures'#13 +
                 ' - Soit parce qu''ils sont utilisés par des frais'#13 +
                 ' - Soit parce qu''ils sont utilisés par des transactions;W;O;O;O;', '', '');

  TFMul(Ecran).BCherche.Click;
end;

initialization
  RegisterClasses([TOF_TRMULFLUXTRESO]);

end.
