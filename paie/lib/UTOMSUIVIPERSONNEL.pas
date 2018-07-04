{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 26/02/2004
Modifié le ... :   /  /
Description .. : TOM gestion du suivi du personnel
Mots clefs ... : PAIE;
*****************************************************************}
{
}
unit UTOMSUIVIPERSONNEL;

interface
uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls, Spin,
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, DBCtrls, Fiche,
  {$ELSE}
  eFiche,
  {$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOM, UTOB, HTB97;

type
  TOM_SUIVIPERSONNEL = class(TOM)
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(stArgument: string); override;
    procedure OnLoadRecord; override;
    procedure OnNewRecord; override;
  private
    TypInt, Interim, TypMess, LeTitre: string;
  end;

implementation
{ TOM_SUIVIPERSONNEL }

procedure TOM_SUIVIPERSONNEL.OnArgument(stArgument: string);
var
  St: string;
begin
  inherited;
  St := Trim(StArgument);
  TypMess := ReadTokenSt(St);
  TypInt := ReadTokenSt(St);
  Interim := ReadTokenSt(St);
  
  SetControlEnabled ('PSV_PGTYPSUIVI', FALSE);
  SetControlEnabled ('PSV_TYPEINTERIM', FALSE);
  SetControlEnabled ('PSV_INTERIMAIRE', FALSe);
  if TypMess = 'EAN' then
  begin
    SetControlVisible ('PSV_OFFREEMPLOI', FALSE) ;
    SetControlVisible ('PSV_REPONSOFFRE', FALSE) ;
    SetControlVisible ('PSV_SALARIE', FALSE) ;
    SetControlVisible ('TPSV_OFFREEMPLOI', FALSE) ;
    SetControlVisible ('TPSV_REPONSOFFRE', FALSE) ;
    SetControlVisible ('TPSV_SALARIE', FALSE) ;
  end;

  LeTitre := 'Suivi de : ' + RechDom('PGTYPSUIVI', TypMess, FALSE);
  LeTitre := LeTitre + ' pour : ' + RechDom('PGINTERIMAIRES', INTERIM, FALSE);
  TFFiche(Ecran).Caption := LeTitre;
  UpdateCaption(Ecran);

end;

procedure TOM_SUIVIPERSONNEL.OnChangeField(F: TField);
begin
  inherited;
end;

procedure TOM_SUIVIPERSONNEL.OnLoadRecord;
begin
  inherited;
end;

procedure TOM_SUIVIPERSONNEL.OnNewRecord;
var
  St : String ;
  Q  : TQuery;
begin
  inherited;
  SetField('PSV_PGTYPSUIVI', TypMess);
  SetField('PSV_TYPEINTERIM', TypInt);
  SetField('PSV_INTERIMAIRE', Interim);
  SetField('PSV_DATEENTRETIEN', NOW) ;
  SetControlEnabled ('PSV_PGTYPSUIVI', FALSE);
  SetControlEnabled ('PSV_TYPEINTERIM', FALSE);
  SetControlEnabled ('PSV_INTERIMAIRE', FALSe);

  st := 'SELECT MAX(PSV_RANG) RANG FROM SUIVIPERSONNEL WHERE PSV_PGTYPSUIVI="'+TypMess+
        '" AND PSV_TYPEINTERIM="'+TypInt+'" AND PSV_INTERIMAIRE="'+Interim+'"';
  Q := OpenSQL(st, True);
  if not Q.eof then SetField('PSV_RANG', Q.FindField('RANG').AsInteger + 1)
  else SetField('PSV_RANG', 1);
  ferme(Q);
  
end;

initialization
  registerclasses([TOM_SUIVIPERSONNEL]);
end.

