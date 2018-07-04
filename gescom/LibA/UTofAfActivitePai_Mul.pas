{***********UNITE*************************************************
Auteur  ...... : TG
Créé le ...... : 21/05/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFACTIVITEPAI_MUL ()
Mots clefs ... : TOF;ACTIVITEPAI_MUL
*****************************************************************}
Unit UTofAfActivitePai_Mul;

Interface

Uses StdCtrls, Controls, Classes,HDB,
     HTB97,
{$IFNDEF EAGLCLIENT}
     db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     FE_Main,
     Mul,
{$ELSE}
     MaineAGL,
     eMul,
{$ENDIF}
     forms, sysutils, ComCtrls,ParamSoc,
     HCtrls, HEnt1, HMsgBox,
     UTOF, UTOFAFTRADUCCHAMPLIBRE ;

Type
  TOF_ACTIVITEPAI_MUL = Class (TOF_AFTRADUCCHAMPLIBRE)
    procedure OnArgument(St : String)  ; override ;
    procedure OnUpdate                 ; override ;
  private
    procedure FlisteDblClick(Sender: TObject);
  end ;

procedure AFLanceFiche_MulActivitePaie;

Implementation


procedure TOF_ACTIVITEPAI_MUL.OnArgument(St: String);
begin
	fMulDeTraitement := true;
  FTableName := 'ACTIVITEPAIE';
  inherited;
{$IFDEF BTP}
  THValComboBox(GetControl('ACP_TYPEARTICLE')).Plus := 'AND (CO_CODE="PRE" OR CO_CODE="FRA")';
{$ENDIF}
	THDbgrid(GetControl('FLISTE')).ondblclick := FlisteDblClick;
end;

procedure TOF_ACTIVITEPAI_MUL.FlisteDblClick(Sender: TObject);
Var StArg 				: String;
    ClePiece			: array [0..7] of Variant;
    TheChaine     : string;
		Fliste	: THDbGrid;
begin
		Fliste := THdbgrid(GetCOntrol('FLISTE'));
  	AglLanceFiche('AFF', 'AFPAIELIENPARAM', '', Fliste.datasource.dataset.FindField('ACP_RANG').asstring, '');
    //
    if GetParamSocSecur('SO_REFRESHMUL', true) then TToolBarButton97 (GetControl('Bcherche')).Click;

end;

procedure TOF_ACTIVITEPAI_MUL.OnUpdate ; // Repompé de UTofAfArticle_Mul (beuaaah)
var F : TFMul;
    i : integer;
    CC : THLabel;
    s : string;
begin
   inherited;
   if not (Ecran is TFMul) then exit;
   F := TFMul(Ecran);

   for i := 1 to 3 do
     with THLabel(GetControl('TACP_FAMILLENIV'+InttoStr(i))) do
       Caption := RechDom('GCLIBFAMILLE', 'LF'+InttoStr(i), false);

{$IFDEF EAGLCLIENT}
   for i := 0 to F.FListe.ColCount-1 do
   begin
      s := F.FListe.Cells[i,0];
      if UpperCase(Copy(s, 1, 7)) = 'FAMILLE' then
      begin
         CC := THLabel(F.FindComponent('TACP_FAMILLENIV'+Copy(s, Length(s), 1)));
         if not CC.Visible then F.Fliste.ColWidths[i] := 0;
         F.Fliste.Cells[i,0] := CC.Caption;
      end;
   end;
{$ELSE}
   for i := 0 to F.FListe.Columns.Count-1 do
   begin
      s := F.FListe.Columns[i].Title.Caption;
      if UpperCase(Copy(s, 1, 7)) = 'FAMILLE' then
      begin
        CC := THLabel(GetControl('TACP_FAMILLENIV'+Copy(s, Length(s), 1)));
        F.Fliste.columns[i].Visible := CC.Visible;
        F.Fliste.columns[i].Field.DisplayLabel := CC.Caption;
      end;
   end;
{$ENDIF}
end;


procedure AFLanceFiche_MulActivitePaie;
begin
   AGLLanceFiche('AFF', 'AFPAIELIEN_MUL', '', '', '');
end;

Initialization
  registerclasses ( [ TOF_ACTIVITEPAI_MUL ] ) ;
end.

