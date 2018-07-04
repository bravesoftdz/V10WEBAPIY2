{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 07/05/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFREVREGULFACTURE ()
Mots clefs ... : TOF;AFREVREGULFACTURE
*****************************************************************}
Unit utofAfRevRegulFacture ;

Interface

Uses StdCtrls, Controls, Classes, Vierge,

{$IFNDEF EAGLCLIENT}
     db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Fe_Main, mul, DBGrids,
{$Else}
     MainEagl, eMul, utob,

{$ENDIF}                     

     TraducAffaire,
     HDB, forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox,
     HQry, UTOF,UtilRevision,UTofAfBaseCodeAffaire, UAFO_FACTREGUL,
     paramsoc;
  
Type
  TOF_AFREVREGULFACTURE = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure BOuvrirOnClick(Sender : TObject);
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;

    private
      {$IFDEF EAGLCLIENT}
      fListe : THGrid;
      {$ELSE}
      fListe : THDBGrid;
      {$ENDIF}
      fDateRegul : TDateTime;

  end;

Type
     TOF_AFREVREGULFAC_VAL = Class (TOF)
     public
     		procedure OnArgument(stArgument : String ) ; override ;
        procedure OnClose ; override ;
        procedure OnUpdate ; override ;
     END ;

procedure AFLanceFicheRegulFacture ;
Implementation

procedure TOF_AFREVREGULFACTURE.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
  Aff1:=THEdit(GetControl('GP_AFFAIRE1'));
  Aff2:=THEdit(GetControl('GP_AFFAIRE2'));
  Aff3:=THEdit(GetControl('GP_AFFAIRE3'));
  Aff4:=THEdit(GetControl('GP_AVENANT'));
end;                          


procedure TOF_AFREVREGULFACTURE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFREVREGULFACTURE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFREVREGULFACTURE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_AFREVREGULFACTURE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AFREVREGULFACTURE.OnArgument (S : String ) ;
var
  i   : Integer;
  vSt : String;

begin
  Inherited;

  {$IFDEF EAGLCLIENT}
    TraduitAFLibGridSt(TFMul(Ecran).FListe);
  {$ELSE}
    TraduitAFLibGridDB(TFMul(Ecran).FListe);
  {$ENDIF}
  fListe := TFMul(Ecran).FListe;

  TFMul(Ecran).BOuvrir.OnClick := BOuvrirOnClick;
  FlagueLesLignesRegularisables(True);
  setcontrolVisible('PCOMPLEMENT', false);
end;

procedure TOF_AFREVREGULFACTURE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_AFREVREGULFACTURE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AFREVREGULFACTURE.OnCancel () ;
begin
  Inherited ;
end ;



procedure TOF_AFREVREGULFACTURE.BOuvrirOnClick(Sender : TObject);
var
  i           : Integer;
  vStWhere    : String;
  MaRegul     : TFACTREGUL;
  vStComment  : String;
  vStDate     : String;
  St          : String;

begin

  vStComment := AGLLanceFiche('AFF','AFREVREGULFAC_VAL','','','');
  vStDate := ReadTokenSt(vStComment);
  if StrToDate(vStDate) = iDate1900 then exit;

  vStWhere := Copy(RecupWhereCritere(TPageControl(GetControl('PAGES'))),6,length(RecupWhereCritere(TPageControl(GetControl('PAGES')))) );

  // on ajoute aux criteres la liste des factures sélectionnées
  try
    if (not fListe.AllSelected) and (fListe.nbSelected <> 0) then
    begin
      for i := 0 to fListe.nbSelected -1 do
        begin
          fListe.GotoLeBookMark(i);
          if (i = 0) and (vStWhere <> '') then
            vStWhere := vStWhere + ' AND ( '
          else if (i <> 0) and (vStWhere <> '') then
            vStWhere := vStWhere + ' OR '
          else
            vStWhere := '(';
          vStWhere := vStWhere + 'GP_NUMERO = ' + TFMul(Ecran).Q.Findfield('GP_NUMERO').AsString;
        end;
      if vStWhere <> '' then vStWhere := vStWhere + ')';
    end;

    St := 'Confirmez vous la validation de ces Factures au ' + vStDate;
    If (PGIAsk(st, ecran.caption) <> mrYes) then exit;

    fDateRegul := strToDate(vStDate);
    MaRegul := TFACTREGUL.Create;
    try
      MaRegul.RegulFactures(vStWhere, fDateRegul, vStComment)
    finally
      MaRegul.Free;
    end;

  // recharger le mul
  FlagueLesLignesRegularisables(True);
  TFMul(Ecran).BCherche.Click;

  except
    on E:Exception do
    begin
      MessageAlerte('Erreur lors du chargement de la liste des factures.' +#10#13#10#13 + E.message );
    end; // on
  end; // try
end;

//*****************************************************************************
//
//   TOF_AFPIECEPRO_VAL  : Saisie date définitive
//
//*****************************************************************************
Function CtrlDate(Zdat : string) : integer;
begin
  result := 0;
  if not(IsValidDate(Zdat)) then result := 1;
end;

procedure TOF_AFREVREGULFAC_VAL.OnArgument(stArgument : String );
Var
  vSt : string;
  i   : Integer;

begin
  Inherited;
  SetControlText('ZZDATE', datetostr(date));
  vSt := GetParamSoc('SO_AFREGULLIB');
  if vSt <> '' then
  begin
    i := Pos ('$$', GetParamSoc('SO_AFREGULLIB'));            // $$  pour reprendre la date
    If (i <> 0) then
      SetControlText('ZZCOMMENTAIRE', Stringreplace(vSt,'$$', datetostr(date), [rfReplaceAll,rfIgnoreCase]));
  end;
END;

procedure TOF_AFREVREGULFAC_VAL.OnUpdate;
Var
  ST  : String;
	Ret : integer;

begin
  inherited ;
  st := GetControlText('ZZDATE');
  ret:= CtrlDate(st);
  if (ret = 1) then PGIInfo('la date saisie n''est pas valide', ecran.Caption);
end;
 
procedure TOF_AFREVREGULFAC_VAL.OnClose;
begin
  inherited;
  if (ctrldate(getcontroltext('ZZDATE')) = 0)  then
    TfVierge(Ecran).retour :=  getcontroltext('ZZDATE') + ';' + getcontroltext('ZZCOMMENTAIRE')
  else
    TfVierge(Ecran).retour :=  '0';
End;

procedure AFLanceFicheRegulFacture ;
begin
  AGLLanceFiche('AFF','AFREVREGULFACTURE','','','');
end ;

Initialization
  registerclasses ( [ TOF_AFREVREGULFACTURE ] ) ;
  registerclasses ( [ TOF_AFREVREGULFAC_VAL ] ) ;
end.
