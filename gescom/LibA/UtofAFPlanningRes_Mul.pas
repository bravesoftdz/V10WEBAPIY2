unit UtofAFPlanningRes_Mul;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,HCtrls,UTOF,
     DicoAF,
      // grc
      HQry,M3FP,HStatus, HMsgBox,
{$IFDEF EAGLCLIENT}
     eMul,Maineagl,
{$ELSE}
     Mul,HDB,FE_Main, dbTables, db,
{$ENDIF}
     Ent1,EntGC,
     TraducAffaire, AFPlanning,
     UtilPGI,HEnt1,UtilGc,
     ConfidentAffaire,UTOFAFTRADUCCHAMPLIBRE,UtilTaches
     ;
Type
     TOF_AFPLANNINGRES_MUL = Class (TOF_AFTRADUCCHAMPLIBRE)
        procedure OnArgument(stArgument : String ) ; override ;
        procedure OnUpdate ; Override;
        procedure ColTauxInvisibles;
        // grc
//        Function RameneListeRessources : String;
        procedure BOuvrirOnClick(Sender : TObject);

        private
{$IFDEF EAGLCLIENT}
          fListe : THGrid;
{$ELSE}
          fListe : THDBGrid;
{$ENDIF}               
     END ;

// grc
//Function AGLRameneListeRessources(parms:array of variant; nb: integer ) : variant ;
Procedure AFLanceFiche_Mul_Ressource;

Function AFLanceFicheAFPlanningRes_Mul(Lequel, Argument : String) : String;

implementation

Function AFLanceFicheAFPlanningRes_Mul(Lequel, Argument : String) : String;
begin
  result := AGLLanceFiche('AFF','AFPLANNINGRES_MUL','', '',Argument);
end;

procedure TOF_AFPLANNINGRES_MUL.OnArgument(stArgument : String );
var      vStTmp, vStChamp, vStValeur : string;
begin

  Inherited;

  UpdateCaption(Ecran);
  setcontrolVisible('BSELECTALL', true);
                       
  {$IFDEF CEGID}
  SetControlVisible('PComplement',False); SetControlVisible('PCompl',False);
  SetControlVisible('BInsert',False);
  {$ENDIF}

  if not AffichageValorisation then
      begin
      SetControlVisible('ARS_TAUXREVIENTUN', false);
      SetControlVisible('ARS_TAUXREVIENTUN_', false);
      SetControlVisible('TARS_TAUXREVIENTUN', false);
      SetControlVisible('TARS_TAUXREVIENTUN_', false);
      end;

  if V_PGI.LaSerie=S3 then TFMul(Ecran).Q.Liste:='RTRESSOURCES';

  // traitement des arguments
  vStTmp:= (Trim(ReadTokenSt(stArgument)));
  While (vStTmp <>'') do
    Begin
      DecodeArgument(vStTmp, vStChamp, vStValeur);
      if vStChamp = 'PLANNING2' then
        Begin
          TFMUL(ecran).DBListe := 'AFMULPLANNING';
//          Ecran.Caption := TraduitGA('Planning des Tâches par Ressource/Affaire');
          Ecran.Caption := TraduitGA('Tâches par Ressource / Affaire');
          UpdateCaption(Ecran);
          setControlText('SELECTION','PLANNING2');
        End
 
      else if vStChamp = 'PLANNING4' then
        Begin
          TFMUL(ecran).DBListe := 'AFMULPLANNING';
//          Ecran.Caption := TraduitGA('Planning des Affaires par Ressource/Tâche');
          Ecran.Caption := TraduitGA('Affaires par Ressource / Tâche');
          UpdateCaption(Ecran);
          setControlText('SELECTION','PLANNING4');
        End

      Else if vStChamp = 'NUMPLANNING' then
        setControlText('NUMPLANNING',vStValeur);

      vStTmp:=(Trim(ReadTokenSt(stArgument)));
    End;

  {$IFDEF EAGLCLIENT}
   TraduitAFLibGridSt(TFMul(Ecran).FListe);
  {$ELSE}
    TraduitAFLibGridDB(TFMul(Ecran).FListe);
  {$ENDIF}

   TFMul(Ecran).BOuvrir.OnClick := BOuvrirOnClick ;
   fListe := TFMul(Ecran).FListe;

   // temporaire
   // le dev n'est pas terminé
   SetControlVisible('CONGES', False);
   {$IFDEF CCS3}
  if (getcontrol('PComplement') <> Nil) then SetControlVisible ('PComplement', true);
  if (getcontrol('PComp') <> Nil) then SetControlVisible ('PComp', true);
  if (getcontrol('PZONE') <> Nil) then SetControlVisible ('PZONE', False);
  {$ENDIF}
      
End;


procedure TOF_AFPLANNINGRES_MUL.ColTauxInvisibles;
{$IFNDEF EAGLCLIENT}
var i : Integer;
{$ENDIF}
Begin
  {$IFNDEF EAGLCLIENT}
  For i:=0 to TFMul(Ecran).FListe.Columns.Count-1 do
    Begin
      If (TFMul(Ecran).FListe.Columns[i].FieldName = 'ARS_TAUXREVIENTUN')
          OR (TFMul(Ecran).FListe.Columns[i].FieldName = 'ARS_PVHT')
          OR (TFMul(Ecran).FListe.Columns[i].FieldName = 'ARS_PVTTC')
          OR (TFMul(Ecran).FListe.Columns[i].FieldName = 'ARS_TAUXUNIT') Then
        TFMul(Ecran).FListe.columns[i].visible:=False ;
    end;
  {$ENDIF}
End;

procedure TOF_AFPLANNINGRES_MUL.OnUpdate;
Begin
if not AffichageValorisation then
    begin
    ColTauxInvisibles;
    end ;

{$IFDEF EAGLCLIENT}
TraduitAFLibGridSt(TFMul(Ecran).FListe);
{$ELSE}
TraduitAFLibGridDB(TFMul(Ecran).FListe);
{$ENDIF}
End;

// grc
{Function AGLRameneListeRessources( parms: array of variant; nb: integer ) : variant;
var  F : TForm ;
     TOTOF  : TOF;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
  if (TOTOF is TOF_AFPLANNINGRES_MUL) then result:=TOF_AFPLANNINGRES_MUL(TOTOF).RameneListeRessources else exit;
end;
}
{Function TOF_AFPLANNINGRES_MUL.RameneListeRessources : String;
var F : TFMul;
    Q : THQuery;
    i : integer;
    code : string;

    {$IFDEF EAGLCLIENT
    L : THGrid;
    {$ELSE
    L : THDBGrid;
    {$ENDIF

begin
  Result := '';
  F:=TFMul(Ecran);
  L:= F.FListe;
  Q:= F.Q;

  if L.AllSelected then
    begin
      Q.First ;
      While not Q.Eof do
        begin
          code:=Q.FindField('ARS_RESSOURCE').asstring ;
          if Result = '' then Result:=code else Result:=Result+';'+code;
          Q.Next ;
        end;
      L.AllSelected:=False;
    end
  else
    if F.FListe.NbSelected <> 0 then
      begin
        InitMove(L.NbSelected,'');
        for i:=0 to L.NbSelected-1 do
          begin
            MoveCur(False);
            L.GotoLeBookmark(i);
            {$IFDEF EAGLCLIENT
            Q.TQ.Seek(L.Row-1) ;
            {$ENDIF
            code:=TFmul(Ecran).Q.FindField('ARS_RESSOURCE').asstring ;
            if Result = '' then Result:=code else Result:=Result+';'+code;
          end;
        L.ClearSelected;
      end;
  FiniMove;
end;
}

procedure TOF_AFPLANNINGRES_MUL.BOuvrirOnClick(Sender : TObject);
var
  i                 : Integer;
  vStWhere          : String;
  vStEtats          : String;
  vStEtat           : String;
  vStWhereEtatLigne : String;
  vStConges         : String;
 
begin

  vStWhere := Copy(RecupWhereCritere(TPageControl(GetControl('PAGES'))),6,length(RecupWhereCritere(TPageControl(GetControl('PAGES')))) );

  // ajout du critere etat des lignes de planning
  if GetControlText('ETATLIGNE') <> '' then
  begin
    i := 0;
    vStEtats := GetControlText('ETATLIGNE');
    While vStEtats <> '' do
    begin
      vStEtat := ReadTokenSt(vStEtats);
      if i = 0 then
        vStWhereEtatLigne := ' AND (APL_ETATLIGNE = "' + vStEtat + '"'
      else
        vStWhereEtatLigne := vStWhereEtatLigne + ' OR APL_ETATLIGNE = "' + vStEtat + '"';
      i := i + 1;
    end;
    vStWhereEtatLigne := vStWhereEtatLigne + ') ';
  end;

  // on ajoute aux criteres la liste des ressources sélectionnées
  try
    if (not fListe.AllSelected) and (fListe.nbSelected <> 0) then
      begin
        for i := 0 to fListe.nbSelected -1 do
          begin
            fListe.GotoLeBookMark(i);
            if (i = 0) and (vStWhere <> '') then vStWhere := vStWhere + ' AND ( '
            else if (i <> 0) and (vStWhere <> '') then vStWhere := vStWhere + ' OR ';
            vStWhere := vStWhere + 'ARS_RESSOURCE = "' + TFMul(Ecran).Q.Findfield('ARS_RESSOURCE').AsSTring + '"';
          end;
        if vStWhere <> '' then vStWhere := vStWhere + ')';
      end;
   except
   on E:Exception do
    begin
     MessageAlerte('Erreur lors du chargement du planning' +#10#13#10#13 + E.message ) ;
    end; // on
   end; // try

  if TCheckBox(GetControl('CONGES')).Checked then
    vStConges := 'X'
  else
    vStConges := '-';

  if (GetControlText('SELECTION') = 'PLANNING2') or
     (GetControlText('SELECTION') = 'PLANNING4') then
    ExecPlanning(GetControlText('NUMPLANNING'), '01/01/1900', vStWhere, vStWhereEtatLigne, vStConges)
                                                                       
end;
 
Procedure AFLanceFiche_Mul_Ressource;
begin
  AGLLanceFiche('AFF','RESSOURCE_Mul','','','');
end;

Initialization
  registerclasses([TOF_AFPLANNINGRES_MUL]);
  //grc         
//  RegisterAglFunc('RameneListeRessources', TRUE , 0, AGLRameneListeRessources);
end.

