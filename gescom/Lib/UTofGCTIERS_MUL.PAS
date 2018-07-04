{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 21/08/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : GCTIERS_MUL ()
Mots clefs ... : TOF;GCTIERS_MUL
*****************************************************************}
Unit UtofGCTIERS_MUL ;

Interface

Uses StdCtrls, Controls, Classes,AGLInit,M3FP,ParamSoc,
{$IFDEF EAGLCLIENT}
      UtileAGL,eFiche,eFichList,maineagl,Spin,emul,
{$ELSE}
      DBCtrls,db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FichList,Fiche,Fe_Main,HDB,Mul,
{$ENDIF}
     forms, sysutils,ComCtrls,
     HCtrls, HTB97, HEnt1, HMsgBox, UTOF,ent1,UtilPgi ;

Type
  TOF_GCTIERS_MUL = Class (TOF)
   procedure OnArgument (Arguments : String ) ;  override ;
   procedure OnUpdate ; override ;
   procedure Onload ; override;
  private

    procedure RenommageEnteteColonnes;

  end ;

Implementation
uses UFonctionsCBP;

procedure TOF_GCTIERS_MUL.OnArgument (Arguments : String ) ;
var stArg, Critere, ChampMul : string;
    PosEgal : integer;
begin
  Inherited ;

  {$IFDEF NOMADE} //Ajoute l'affichage des prospects dans le mul
  SetControlText('XX_WHERE','OR T_NATUREAUXI="PRO"');
  {$ELSE}
  SetControlText('XX_WHERE','');
  {$ENDIF}

  SetControlText('TYPEACTION','ACTION=MODIFICATION'); //Initialisation à modification
  stArg:=Arguments; //Cache les boutons ajout et dupliquer si consultation
  Repeat
    Critere := UpperCase(Trim(ReadTokenSt(stArg)));
    if Critere<>'' then
      begin
      PosEgal := pos('=',Critere);
      if PosEgal<>0 then ChampMul := copy(Critere,1,PosEgal-1);
        if ChampMul='ACTION' then SetControlText('TYPEACTION',Critere);
      end;
  until Critere='';
  SetControlVisible('bInsert',StringToAction(GetControlText('TYPEACTION'))<>taConsult);
  SetControlVisible('B_DUPLICATION',StringToAction(GetControlText('TYPEACTION'))<>taConsult);

  if Not ExJaiLeDroitConcept(TConcept(gcCLICreat),False) then
  BEGIN
    SetControlVisible('BINSERT',False) ;
    SetControlVisible('B_DUPLICATION',False) ;
  END ;

end ;

Function TOF_GCTIERS_MUL_AGLNatureTiers (Parms: array of variant; nb: integer) : variant ;
begin
{$IFDEF NOMADE}
if ctxMode in V_PGI.PGIContexte then Result:='CLI'
else if GetParamSoc('SO_GCPREFIXETIERS')<>'' then Result:='PRO' else Result:='0';
{$ELSE}
Result:='CLI';
{$ENDIF}
end;

procedure TOF_GCTIERS_MUL.Onload;
begin
  inherited;
  RenommageEnteteColonnes;
end;

procedure TOF_GCTIERS_MUL.OnUpdate;
begin
  inherited;
  RenommageEnteteColonnes;
end;

procedure TOF_GCTIERS_MUL.RenommageEnteteColonnes;
var i : integer;
		Gr : THDbgrid;
    stChamp,Libelle : string;
begin
	Gr := TFMul(ecran).fliste;
	For i:=0 to Gr.Columns.Count-1 do
  Begin
    StChamp := TFMul(Ecran).Q.FormuleQ.GetFormule(Gr.Columns[i].FieldName);
    if copy(UpperCase (stChamp),1,14)='YTC_TABLELIBRE' then
    begin
      libelle := RechDomZoneLibre('CT'+Copy(stChamp,15,1),false);
{$IFNDEF AGL581153}
			TFMul(ecran).SetDisplayLabel (StChamp,TraduireMemoire(Libelle));
{$else}
			TFMul(ecran).SetDisplayLabel (i,TraduireMemoire(Libelle));
{$endif}
    end else  if copy(UpperCase (stChamp),1,13)='YTC_BOOLLIBRE' then
    begin
      libelle := RechDomZoneLibre('CB'+Copy(stChamp,14,1),false);
{$IFNDEF AGL581153}
			TFMul(ecran).SetDisplayLabel (StChamp,TraduireMemoire(Libelle));
{$else}
			TFMul(ecran).SetDisplayLabel (i,TraduireMemoire(Libelle));
{$endif}
    end else  if copy(UpperCase (stChamp),1,13)='YTC_DATELIBRE' then
    begin
      libelle := RechDomZoneLibre('CD'+Copy(stChamp,14,1),false);
{$IFNDEF AGL581153}
			TFMul(ecran).SetDisplayLabel (StChamp,TraduireMemoire(Libelle));
{$else}
			TFMul(ecran).SetDisplayLabel (i,TraduireMemoire(Libelle));
{$endif}
    end else  if copy(UpperCase (stChamp),1,14)='YTC_TEXTELIBRE' then
    begin
      libelle := RechDomZoneLibre('CC'+Copy(stChamp,15,1),false);
{$IFNDEF AGL581153}
			TFMul(ecran).SetDisplayLabel (StChamp,TraduireMemoire(Libelle));
{$else}
			TFMul(ecran).SetDisplayLabel (i,TraduireMemoire(Libelle));
{$endif}
    end;

  end;
end;

Initialization
  registerclasses ( [TOF_GCTIERS_MUL] ) ;
  RegisterAglFunc('TiersNature',True,1,TOF_GCTIERS_MUL_AGLNatureTiers);
end.
