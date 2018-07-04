{***********UNITE*************************************************
Auteur  ...... : FC
Créé le ...... : 25/04/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : UTOFPGMUL_EMETTEURSOCIAL ()
Mots clefs ... : TOF;UTOFPGMUL_EMETTEURSOCIAL
*****************************************************************}
Unit UTOFPGMUL_EMETTEURSOCIAL ;

Interface

Uses StdCtrls, 
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, HDB,
{$else}
     eMul, 
     uTob,UtileAGL,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF,
     HTB97 ;

Type
  TOF_UTOFPGMUL_EMETTEURSOCIAL = Class (TOF)
    public
      procedure OnArgument (S : String ) ; override ;
    private
      procedure DeleteClick (Sender: TObject);
  end ;

Implementation

procedure TOF_UTOFPGMUL_EMETTEURSOCIAL.OnArgument (S : String ) ;
var
  Delete : TToolbarButton97;
begin
  Inherited ;
  Delete := TToolbarButton97(GetControl('BDELETE'));
  if (Delete <> nil) then
    Delete.OnClick:= DeleteClick;
end ;

procedure TOF_UTOFPGMUL_EMETTEURSOCIAL.DeleteClick(Sender: TObject);
var
  {$IFNDEF EAGLCLIENT}
    Liste : THDBGrid;
  {$ELSE}
    Liste : THGrid;
  {$ENDIF}
  reponse,i : integer;
  BCherche : TToolbarButton97;
begin
  {$IFNDEF EAGLCLIENT}
    Liste := THDBGrid(GetControl('FListe'));
  {$ELSE}
    Liste := THGrid(GetControl('FListe'));
  {$ENDIF}

  if (Liste <> nil) then
  begin
    if (Liste.NbSelected = 0) then
    begin
      MessageAlerte('Aucun émetteur social n''est sélectionné');
      exit;
    end;
  end;

  reponse:= PGIAsk ('Cette commande supprimera les émetteurs sociaux sélectionnés.#13#10'+
                  'Voulez-vous continuer ?', TFMul(Ecran).Caption);
  if (reponse <> mrYes) then
    exit;

  try
    BeginTrans;
    for i := 0 to Liste.NbSelected - 1 do
    begin
      Liste.GotoLeBOOKMARK(i);
    {$IFDEF EAGLCLIENT}
      TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row-1) ;
    {$ENDIF}
    ExecuteSQL('DELETE FROM EMETTEURSOCIAL WHERE '+
          'PET_EMETTSOC = "'+ TFmul(Ecran).Q.FindField('PET_EMETTSOC').asstring +'"');
    end;
    Committrans;
  except
    Rollback;
    PGIBox('! Erreur lors de la suppression des émetteurs sociaux sélectionnés','');
  end;

  BCherche:= TToolbarButton97 (GetControl ('BCherche'));
  if BCherche <> nil then
    BCherche.click;
end;

Initialization
  registerclasses ( [ TOF_UTOFPGMUL_EMETTEURSOCIAL ] ) ; 
end.
