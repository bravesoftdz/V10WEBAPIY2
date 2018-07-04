{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 07/10/2009
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTEPUREAFF_MUL ()
Mots clefs ... : TOF;BTEPUREAFF_MUL
*****************************************************************}
Unit BTEPUREAFF_MUL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Ent1,
     EntGc,
     UEntCommun,
     AglInit,
     M3FP,
{$IFNDEF EAGLCLIENT}
		 FE_Main,
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
		 MainEagl,
     eMul,
     HQry,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     utofAfBaseCodeAffaire,
     RECALCPIECE_RAP_TOF,
     DateUtils,
     Paramsoc,
     UTOF ;

Type
  TOF_BTEPUREAFF_MUL = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	TOBChantiers : TOB;
  	procedure EpurelesChantiers;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
    procedure AddChantierAtraiter (Chantier : string);
    function ControleDates : boolean;
  end ;

Implementation

uses TntStdCtrls;

procedure AglBTEpureChantier (parms:array of variant; nb: integer ) ;
var  F : TForm ;
     LaTof : TOF;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFMul) then Latof:=TFMul(F).Latof else exit;
  if (LaTof is TOF_BTEPUREAFF_MUL) then TOF_BTEPUREAFF_MUL(LaTof).EpurelesChantiers else exit;
end;

procedure TOF_BTEPUREAFF_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTEPUREAFF_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTEPUREAFF_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTEPUREAFF_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTEPUREAFF_MUL.OnArgument (S : String ) ;
Var CC : THValComboBox;
begin
  fMulDeTraitement  := true;
  Inherited;
  fTableName := 'AFFAIRE';
  TOBChantiers := TOB.Create ('LES CHANTIERS',nil,-1);

  //Gestion Restriction Domaine et Etablissements
  CC:=THValComboBox(GetControl('AFF_DOMAINE')) ;
  if CC<>Nil then PositionneDomaineUser(CC) ;

  CC:=THValComboBox(GetControl('AFF_ETABLISSEMENT')) ;
  if CC<>Nil then PositionneEtabUser(CC) ;
  //
end ;

procedure TOF_BTEPUREAFF_MUL.OnClose ;
begin
  TOBChantiers.free;
  Inherited ;
end ;

procedure TOF_BTEPUREAFF_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTEPUREAFF_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTEPUREAFF_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2,
  Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers,
  Tiers_: THEdit);
begin
  Aff:=THEdit(GetControl('AFF_AFFAIRE'));
  // MODIF LS
  Aff0 := THEdit(GetControl('AFF_AFFAIRE0'));
  // --
  Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
  Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
  Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
  Aff4:=THEdit(GetControl('AFF_AVENANT'));
  Tiers:=THEdit(GetControl('AFF_TIERS'));
  // affaire de référence pour recherche
  Aff_:=THEdit(GetControl('AFF_AFFAIREREF'));
  Aff1_:=THEdit(GetControl('AFFAIREREF1'));
  Aff2_:=THEdit(GetControl('AFFAIREREF2'));
  Aff3_:=THEdit(GetControl('AFFAIREREF3'));
  Aff4_:=THEdit(GetControl('AFFAIREREF4'));
end;


procedure TOF_BTEPUREAFF_MUL.EpurelesChantiers;
var
{$IFDEF EAGLCLIENT}
		Q : THquery;
{$ELSE}
		Q : Tquery;
{$ENDIF}
		F : TFmul;
    i : integer;
    SRet : String;
    TOBPass : TOB;
    Affaire : string;
    GetOUt,AvecControle : boolean;
begin
  //
  AvecControle := True;
  TOBChantiers.clearDetail;
  GetOut := false;

  if not ControleDates then
  begin
    if PGIAsk('ATTENTION : Vous demandez une épuration en dehors du délai légal de tenue.#13#10 Confirmez-vous le traitement')<>Mryes then
    begin
      Exit;
    end;
  end;

  if GetControl('CBSANSCONTROLE')<> nil then
  begin
    AvecControle := not THCheckbox (GetControl('CBSANSCONTROLE')).Checked;
	    if not AvecControle then
    begin
      if PGIAsk('ATTENTION : Vous demandez une épuration sans contrôle.#13#10 Confirmez-vous le traitement')<>Mryes then
      begin
        Exit;
      end;
    end;
  end;

  TOBPass := TOB.create ('UNE TOB',nil,-1);
  TOBPASS.AddChampSupValeur ('RESULTAT','');
  TheTOB := TOBPAss;
  AGLLAnceFiche('BTP','BTCONFIRMEPURCHA','','','');
  TheTOB := nil;
  if TOBPass.getValue('RESULTAT')<>'OK' then GetOUt := true;
  TOBPass.free;
  //
  if GetOUt then exit;   // pas d'accord
  //
  TOBPass := TOB.create ('UNE TOB',nil,-1);
  TOBPASS.AddChampSupValeur ('DAYPASS','');
  TOBPASS.AddChampSupValeur ('RESULTAT','');
  TheTOB := TOBPAss;
 	Sret := AGLLAnceFiche('BTP','BTCONFIRMPASS','','','');
  TheTOB := nil;
  if TOBPass.GetValue('RESULTAT')<> 'OK' then GetOut := true;
  TOBPass.free;
  //
  if GetOUt then exit; // toujours pas d'accord
  //
  TOBChantiers.AddChampSupValeur ('AVECCONTROLE',BoolToStr_(AvecControle) );

  F := TFmul(ecran);
  TRY
    if TFMul(F).Fliste.AllSelected then
    BEGIN
      Q:=TFmul(F).Q;
      Q.First;
      while Not Q.EOF do
      BEGIN
        Affaire := Q.FindField('AFF_AFFAIRE').AsString;
        AddChantierAtraiter (Affaire);
        Q.NEXT;
      END;
      if TOBChantiers.detail.count > 0 then
      begin
      	TraiteEpureChantiers (TOBChantiers);
      end;
      TFMul(F).Fliste.AllSelected:=False;
    END ELSE
    BEGIN
      for i:=0 to TFMul(F).Fliste.nbSelected-1 do
      begin
        TFMul(F).Fliste.GotoLeBookmark(i);
{$IFDEF EAGLCLIENT}
      	TFMul(F).Q.TQ.Seek(TFMul(F).FListe.Row-1) ;

        Affaire :=TFMul(F).Q.FindField('AFF_AFFAIRE').AsString;
        AddChantierAtraiter (Affaire);
{$ELSE}
        Affaire :=TFMul(F).Fliste.datasource.dataset.FindField('AFF_AFFAIRE').AsString;
        AddChantierAtraiter (Affaire);
{$ENDIF}
      end;
    	if TOBChantiers.detail.count > 0 then
      begin
      	TraiteEpureChantiers (TOBChantiers);
      end;
    END;
  finally
  	TOBChantiers.ClearDetail;
  end;
end;
procedure TOF_BTEPUREAFF_MUL.AddChantierAtraiter(Chantier: string);
var TOBC : TOB;
begin
	TOBC := TOB.Create ('UN CHANTIER',TOBChantiers,-1);
  TOBC.AddChampSupValeur('AFF_AFFAIRE',chantier);
end;

function TOF_BTEPUREAFF_MUL.ControleDates: boolean;

	function EnDate (UnParam : string) :TDateTime;
  begin
    TRY
  		Result := StrToDate(unParam);
    except
      Result := StrToDate('31/12/2099');
    end;
  end;

var MaxDate : TdateTime;
		Delai : Integer;
begin
  Result := false;
  Delai := GetParamSocSecur ('SO_BTDELAILEGALEPU',0);
  if Delai = 0 then begin result := true; Exit; end;
  Delai := Delai * (-1);
	MaxDate := EndOfAMonth(YearOf(Now),MonthOf(Now));
  MaxDate := IncYear(MaxDate,Delai);
  MaxDate := StrToDate(DateToStr(MaxDate)); // les heures etc.. osef
  //
  if (EnDate(GetControlText('AFF_DATEDEBUT')) >= MaxDate) or
  	 (EnDate(GetControlText('AFF_DATEDEBUT_')) >= MaxDate) then Exit;

  if (EnDate(GetControlText('AFF_DATEFIN')) >= MaxDate) or
  	 (EnDate(GetControlText('AFF_DATEFIN_')) >= MaxDate) then Exit;

  Result := True;
end;

Initialization
  registerclasses ( [ TOF_BTEPUREAFF_MUL ] ) ;
	RegisterAglProc('BTEpureChantier', True , 0, AGlBTEpureChantier);
end.
