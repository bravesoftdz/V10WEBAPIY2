Unit UTOFAFREVCALCCOEF ;

Interface

uses  StdCtrls,Controls,Classes,
{$IFDEF EAGLCLIENT}
      eMul,  Maineagl,
{$ELSE}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}mul,FE_Main,
{$ENDIF}
      forms,sysutils,GCMZSUtil,Utob,FactUtil, Ent1,
      ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF, Windows, ParamSoc,
      Dicobtp,SaisUtil,EntGC,M3FP,  UtilMultrt,Hstatus,AffaireUtil,
      utofAfBaseCodeAffaire,utilpgi,AglInit,UtilGc,TraducAffaire,
      AffEcheanceUtil, HTB97, Ed_tools, UAFO_REVPRIXCALCULCOEF,
      UtilRevision;     
                   
Type
  TOF_AFREVCALCCOEF = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
    private
      procedure CalculCoefEnSerie(Sender: TObject);
  end;

const
	TexteMsg : array[1..2] of string 	= (
          {1}        'Affaire : %s. La formule %s n''a pas pu être calculée.',
          {2}        'Affaire : %s. La formule %s n''est pas paramètrée.'
          );
 
Procedure AFLanceFicheRevCalcCoef;

Implementation


procedure TOF_AFREVCALCCOEF.OnArgument (S : String ) ;
begin
  Inherited ;
  TToolBarButton97(GetControl('BOuvrir')).OnClick := CalculCoefEnSerie;
end ;

procedure TOF_AFREVCALCCOEF.OnClose ;
begin
  Inherited;
end;

procedure TOF_AFREVCALCCOEF.CalculCoefEnSerie(Sender: TObject);
Var            
  F             : TFMul ;
  ret           : T_ErrTraiteEnreg;
  vTob          : Tob;
  i             : Integer;
  vBoResult     : Boolean;
  vRev          : TCALCULCOEF;
  vSLLog        : TStringList;
  vTobForm      : Tob;
  vTobLaFormule : Tob;
  vQr           : TQuery;

begin

  vBoResult := True;

  vTob := Tob.Create('AFFAIRE',NIL,-1);
  vTobForm := Tob.create('AFPARAMFORMULE', NIL, -1);
  vQr := nil;
  Try
    vQR := OpenSql('SELECT * FROM AFPARAMFORMULE', True);
    if Not vQR.Eof then
      vTobForm.LoadDetailDB('AFPARAMFORMULE','','',vQR,False,True);

    F := TFMul(Ecran);
    if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
    begin
      MessageAlerte('Aucun élément sélectionné');
      exit;
    end;

    // traitement des affaires
//    ret:=TraiteEnregMulTable (F, 'SELECT * FROM AFPARAMFORMULE', 'AFC_AFFAIRE' , 'AFPARAMFORMULE', 'AFC_AFFAIRE','AFPARAMFORMULE', vTob, True);
    ret := TraiteEnregMulListe (TFMul(Ecran), 'AFF_AFFAIRE','AFAFFTIERS ', vTob, True);
    If ret <> tteOK then exit ;

    InitMoveProgressForm(nil, 'Traitement en cours...', '', vTob.detail.count, false, true);
    Try
      for i := 0 to vTob.Detail.count -1 do
      begin
        MoveCurProgressForm('');

        vTobLaFormule := vTobForm.FindFirst(['AFC_AFFAIRE'], [vTob.Detail[i].GetValue('AFF_AFFAIRE')], True);
        if vTobLaFormule = nil then
        begin
          vBoResult := False;
          DebutLog('', vSLLog);
          vSLLog.Add(format(TexteMsg[2], [DateToStr(now),
                                          vTob.Detail[i].GetValue('AFP_AFFAIRE'),
                                          vTob.Detail[i].GetValue('AFP_FORCODE')]));
          FinLog('', vSLLog);
        end;

        while vTobLaFormule <> nil do
        begin
          vRev := TCALCULCOEF.Create;
          try
            if not vRev.CalculFormuleAuto(vTobLaFormule) then
            begin
              vBoResult := False;
              DebutLog('', vSLLog);
              vSLLog.Add(format(TexteMsg[1], [DateToStr(now),
                                              vTob.Detail[i].GetValue('AFP_AFFAIRE'),
                                              vTob.Detail[i].GetValue('AFP_FORCODE')]));
              FinLog('', vSLLog);
            end;
          finally
            vRev.Free;
          end;
          vTobLaFormule := vTobForm.FindNext(['AFC_AFFAIRE'], [vTob.Detail[i].GetValue('AFF_AFFAIRE')], True);
        end;
      end;
    finally
      FiniMoveProgressForm;
    end;

  if not vBoResult then
    PGIInfoAF('Un ou plusieurs coefficients n''ont pas pu être Calculés. Vous pouvez consulter le fichier .log.', Ecran.caption)
  else
    PGIInfoAF('Traitement terminé.', Ecran.caption);

  finally
    vTob.Free;
    vTobForm.Free;
    ferme(vQr);
  end;
end;

// Attention, pas de AFFAIRE0 déclaré : le champ n'existe pas sur la form
procedure TOF_AFREVCALCCOEF.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
  Aff:=THEdit(GetControl('AFF_AFFAIRE'));
  Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
  Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
  Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
  Aff4:=THEdit(GetControl('AFF_AVENANT'));
  Tiers:=THEdit(GetControl('AFF_TIERS'));
end;

procedure AFLanceFicheRevCalcCoef;
begin
  AGLLanceFiche('AFF','AFREVCALCCOEF','','','');
end;

Initialization
  registerclasses ([TOF_AFREVCALCCOEF]);

end.
