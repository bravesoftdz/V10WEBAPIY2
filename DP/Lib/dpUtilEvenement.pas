unit dpUtilEvenement;
// #### REVOIR TOUT CE SOURCE QUI SEMBLE TORTURE POUR PAS GRAND CHOSE...

interface

uses  StdCtrls,Controls,Classes,db,forms,sysutils,dbTables,ComCtrls,
      extctrls,HCtrls,HEnt1,HMsgBox,UTOM,DBCtrls ,HRichEdt, HRichOLE,
      HStatus,HDB, PGIEnv;

Type
  TEvtSup = Class
    Code : string;
    prefixe : string;
    table   : string;
    Typedos : string;
    Fonction :string;
    Codeper :string;
  public
    procedure SuppressionEnregControle();
    procedure SuppressionEnregControlelien();
  end;
procedure SupprimeListeEnreg (L : THDBGrid; Q :TQuery; St:string);
function SuppressionEvenement(Code : string; Typedos : string; Fonction :string; Codeper :string; prefixe : string; table : string) : boolean;


///////////// IMPLEMENTATION //////////////
implementation

// suppression pour tables contenant un NODP et un NOORDRE
// (contrôles fiscaux, évolution capital)
procedure TEvtSup.SuppressionEnregControle();
var
    res : integer;
    sWhere: String;
    lib_ordre : string;
    lib_cle   : string;
begin
  sWhere := '';
  lib_ordre := prefixe + '_NOORDRE';
  lib_cle   := prefixe + '_NODP';
  sWhere := table + ' WHERE ' + lib_cle + '= ' + (InttoStr(V_PGI_Env.CodePer))+ ' AND '
  + lib_ordre + '= "'+ Code + '"';
  res := ExecuteSQL ('DELETE FROM '+sWhere);
  if res=-1 then V_PGI.IoError:=oeUnknown ;
end;

// supression d'un enregistrement annulien
procedure TEvtSup.SuppressionEnregControlelien();
var
    res : integer;
    sWhere: String;
    cle : string;
begin

  sWhere := '';
  if (table = 'ANNULIEN') then
    sWhere := table + ' Where ANL_CODEPERDOS='+ Code+ ' AND ANL_TYPEDOS="'+Typedos+'"' +' AND ANL_NOORDRE=1 AND ANL_FONCTION="'+Fonction +'"'+' AND ANL_CODEPER='+ Codeper
  else
    begin
    if (table = 'JUEVENEMENT') then
      cle   := prefixe + '_NOEVT'
    else
      cle   := prefixe + '_CODEPER';
    sWhere := table + ' Where '+ cle +'='+Code;
    end;
  res := ExecuteSQL ('DELETE FROM '+sWhere);
  if res=-1 then V_PGI.IoError:=oeUnknown ;
end;


// suppression d'une liste de liens (capital, contrôles fiscaux...)
procedure SupprimeListeEnreg(L : THDBGrid; Q :TQuery; St:string);
var i : integer;
    Texte : string;
begin
  if (L.NbSelected=0) and (Not L.AllSelected) then
  begin
    PGIInfo('Aucun élément sélectionné', TitreHalley);
    exit;
  end;
  Texte:='Vous allez supprimer définitivement les informations.#13#10Confirmez vous l''opération ?';
  if HShowMessage('0;Suppression historique;'+Texte+';Q;YN;N;N;','','')<>mrYes then exit ;
  if L.AllSelected then
    BEGIN
    Q.First;
    while Not Q.EOF do
      begin
      MoveCur(False);
      if St ='EVOLUCAPITAL' then
        begin
        if not SuppressionEvenement(Q.FindField('DPM_NOORDRE').AsString, '', '', '', 'DPM', 'DPMVTCAP') then Break ;
        end;
      if St ='CONTROLEFISC' then
        begin
        if not SuppressionEvenement(Q.FindField('DCL_NOORDRE').AsString, '', '', '', 'DCL', 'DPCONTROLE') then Break ;
        end;
      if St ='ANNULIEN' then
        begin
        if not SuppressionEvenement(Q.FindField('ANL_CODEPERDOS').AsString, Q.FindField('ANL_TYPEDOS').AsString, Q.FindField('ANL_FONCTION').AsString, Q.FindField('ANL_CODEPER').AsString, 'ANL', 'ANNULIEN') then Break ;
        end;
      if St = 'EVENEMENT' then
        begin
        if not SuppressionEvenement(Q.FindField('JEV_NOEVT').AsString, '', '', '', 'JEV', 'JUEVENEMENT') then Break ;
        end;
      // #### géré par AGLSupprimeListAnnu
      if St ='ANNUAIRE' then
        begin
        if not SuppressionEvenement(Q.FindField('ANN_CODEPER').AsString, '', '', '', 'ANN', 'ANNUAIRE') then Break ;
        end;
      Q.Next;
      end;
    L.AllSelected:=False;
    END
  ELSE
    BEGIN
    InitMove(L.NbSelected,'');
    for i:=0 to L.NbSelected-1 do
      begin
      MoveCur(False);
      L.GotoLeBookmark(i);
      if St ='EVOLUCAPITAL' then
        begin
        if not SuppressionEvenement(Q.FindField('DPM_NOORDRE').AsString, '', '', '', 'DPM', 'DPMVTCAP') then Break ;
        end;
      if St ='CONTROLEFISC' then
        begin
        if not SuppressionEvenement(Q.FindField('DCL_NOORDRE').AsString, '', '', '', 'DCL', 'DPCONTROLE') then Break ;
        end;
      if St ='ANNULIEN' then
        begin
        if not SuppressionEvenement(Q.FindField('ANL_CODEPERDOS').AsString, Q.FindField('ANL_TYPEDOS').AsString, Q.FindField('ANL_FONCTION').AsString, Q.FindField('ANL_CODEPER').AsString, 'ANL', 'ANNULIEN') then Break ;
        end;
      if St = 'EVENEMENT' then
        begin
        if not SuppressionEvenement(Q.FindField('JEV_NOEVT').AsString, '', '', '', 'JEV', 'JUEVENEMENT') then Break ;
        end;
      if St ='ANNUAIRE' then
        begin
        if not SuppressionEvenement(Q.FindField('ANN_CODEPER').AsString, '', '', '', 'ANN', 'ANNUAIRE') then Break ;
        end;
      end;
    L.ClearSelected;
    FiniMove;
    END;
end;


function SuppressionEvenement(Code : string; Typedos : string; Fonction :string; Codeper :string; prefixe : string; table : string) : boolean;
var
    EvtSup : TEvtSup;
begin
  EvtSup := TEvtSup.Create;
  EvtSup.Code := Code;
  EvtSup.prefixe := prefixe;
  EvtSup.table := table;
  EvtSup.Typedos := Typedos;
  EvtSup.Fonction := Fonction;
  EvtSup.Codeper := Codeper;
  if (table = 'ANNULIEN') or (table = 'ANNUAIRE')
  or (table = 'JUEVENEMENT') then
    BEGIN
    if Transactions(EvtSup.SuppressionEnregControlelien, 3)<>oeOk then
      begin
      Result := false;
      PGIInfo('Suppression impossible', TitreHalley) ;
      exit;
      end ;
    END
  else
    BEGIN
    if Transactions(EvtSup.SuppressionEnregControle, 3)<>oeOk then
      begin
      Result := false;
      PGIInfo('Suppression impossible', TitreHalley) ;
      exit;
      end ;
    END;
  EvtSup.Free;
  Result := true;
end;


end.
