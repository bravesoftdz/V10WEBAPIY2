{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 24/04/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULINITHISTO ()
Mots clefs ... : TOF;PGMULINITHISTO
*****************************************************************}
Unit UTofPGMulInitHisto ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db, HDB,Fe_Main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,MainEAGL,
{$ENDIF}
     forms,
     sysutils,
     uTob,
     ComCtrls,
     ParamSoc,
     HCtrls, 
     HEnt1,
     Ed_Tools,
     HQry,
     Vierge,
     HMsgBox, 
     UTOF ; 

Type
  TOF_PGMULINITHISTO = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    {$IFNDEF EAGLCLIENT}
    Liste : THDBGrid ;
    {$ELSE}
    Liste : THGrid ;
    {$ENDIF}
    Provenance : String;
    procedure GrilleDblClick(Sender : TObject);
    procedure RecupHisto(Champ : String);
    procedure HistoParDefaut(Champ : String;PourTous : Boolean);
  end ;
  TOF_PGCONFIRMPAIE = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose; override ;
  end;

Implementation

procedure TOF_PGMULINITHISTO.OnLoad ;
var StWhere,StPredef : String;
begin
Inherited ;

  StWhere := '';
  If GetControltext('ETATINIT') = 'OUI' then StWhere := 'PPP_PGINFOSMODIF IN (SELECT DISTINCT(PHD_PGINFOSMODIF) FROM PGHISTODETAIL)'
  else If GetControltext('ETATINIT') = 'NON' then StWhere := 'PPP_PGINFOSMODIF NOT IN (SELECT DISTINCT(PHD_PGINFOSMODIF) FROM PGHISTODETAIL)';
  StPredef := 'PPP_PREDEFINI<>"CEG" AND (PPP_PREDEFINI="STD" OR (PPP_PREDEFINI="DOS" AND PPP_NODOSSIER="'+V_PGI.NoDossier+'")) AND PPP_HISTORIQUE="X"';
  if StWhere <> '' then StWhere := StWhere+ ' AND '+StPredef
  else StWhere := StPredef;
  SetControlText('XX_WHERE',StWhere);
end;

procedure TOF_PGMULINITHISTO.OnArgument (S : String ) ;
begin
  Inherited ;
  Provenance := readTokenPipe(S,';');
  SetControlText('PPP_PGTYPEINFOLS','SAL');
  {$IFNDEF EAGLCLIENT}
  Liste := THDBGrid(GetControl('FListe'));
 {$ELSE}
  Liste := THGrid(GetControl('FListe'));
  {$ENDIF}
  If Liste <> Nil Then Liste.OnDblClick := GrilleDblClick ;
  SetControlText('ETATINIT','NON');
  If GetParamSocSecur('SO_PGHISTORISATION',True) then SetControlText('METHODEINIT','RECUP')
  else
  begin
    SetControlText('METHODEINIT','SALARIE');
    SetControlEnabled('METHODEINIT',False);
  end;
end ;

procedure TOF_PGMULINITHISTO.GrilleDblClick(Sender : TObject);
var Q_Mul : THQuery;
    Champ,StSal : String;
    i : Integer;
    ListeAvertissement : String;
    Rep : Word;
    Retour : String;
begin
  Q_Mul := THQuery(Ecran.FindComponent('Q'));
  {$IFNDEF EAGLCLIENT}
  Liste := THDBGrid(GetControl('FListe'));
  {$ELSE}
  Liste := THGrid(GetControl('FListe'));
  {$ENDIF}
  If provenance <> 'PARAMSOC' then
  begin
    If (not Liste.AllSelected) AND (Liste.NbSelected = 0)then
    begin
      PGIBox('Aucun élément selectionné',Ecran.Caption);
      Exit;
    end;
  end;
  //Verification des champs pour affiche rmessage d'avertissement
  ListeAvertissement := '';
  If Liste.AllSelected then
  begin
    {$IFDEF EAGLCLIENT}
    if (TFMul(Ecran).bSelectAll.Down) then
    TFMul(Ecran).Fetchlestous;
    {$ENDIF}
    Q_Mul.First;
    While not Q_Mul.eof do
    begin
      Champ := Q_Mul.FindField('PPP_PGINFOSMODIF').AsString;
      If ExisteSQL('SELECT PHD_SALARIE FROM PGHISTODETAIL WHERE PHD_PGINFOSMODIF="'+Champ+'"') then
      begin
        ListeAvertissement := ListeAvertissement + RechDom('PGCHAMPSAL',Champ,False);
      end;
      Q_Mul.Next;
    end;
  end
  else
  begin
    for i := 0 to Liste.NbSelected-1 do
    begin
      Liste.GotoLeBookmark(i);
      {$IFDEF EAGLCLIENT}
      TFmul(Ecran).Q.TQ.Seek(Liste.Row-1) ;
      {$ENDIF}
      Champ := Q_Mul.FindField('PPP_PGINFOSMODIF').AsString;
      If ExisteSQL('SELECT PHD_SALARIE FROM PGHISTODETAIL WHERE PHD_PGINFOSMODIF="'+Champ+'"') then
      begin
        ListeAvertissement := ListeAvertissement + '#13#10 - '+RechDom('PGCHAMPSAL',Champ,False);
      end;
    end;
  end;
  If ListeAvertissement <> '' then
  begin
    Rep := PGIAsk('Attention l''historique contient déja des données pour les informations suivantes : '+ListeAvertissement+
    '#13#10 L''initialisation va détruire le contenu actuel, voulez vous initialiser l''historique ?',
          ecran.Caption);
        if Rep = MrYes then
        begin
           Retour := AglLanceFiche('PAY', 'CONFIRMPAIE', '', '', 'Etes vous certain de vouloir réinitialiser l''historique ?');
           if retour <> 'OUI' then
           begin
            Exit;
           end;
        end
        else Exit;
  end;


  If TFMul(Ecran).bSelectAll.Down then Liste.AllSelected := True;
  If (Liste.AllSelected) or (Provenance='PARAMSOC') then
  begin
    {$IFDEF EAGLCLIENT}
    if (TFMul(Ecran).bSelectAll.Down) or (Provenance='PARAMSOC') then
    TFMul(Ecran).Fetchlestous;
    {$ENDIF}
    InitMoveProgressForm (NIL,'Initialisation de l''historique ',
                    'Veuillez patienter SVP ...', Q_Mul.RecordCount,
                    False,True);
    Q_Mul.First;
    While not Q_Mul.eof do
    begin
      Champ := Q_Mul.FindField('PPP_PGINFOSMODIF').AsString;
      If GetControlText('METHODEINIT') = 'RECUP' then RecupHisto(Champ)
      else HistoParDefaut(Champ,True);
      MoveCurProgressForm ('Information récupérée : '+RechDom('PGCHAMPSAL',Champ,False));
      Q_Mul.Next;
    end;
  end
  else
  begin
    InitMoveProgressForm (NIL,'Initialisation de l''historique ',
                    'Veuillez patienter SVP ...', Liste.NbSelected,
                    False,True);
    for i := 0 to Liste.NbSelected-1 do
    begin
      Liste.GotoLeBookmark(i);
      {$IFDEF EAGLCLIENT}
      TFmul(Ecran).Q.TQ.Seek(Liste.Row-1) ;
      {$ENDIF}
      Champ := Q_Mul.FindField('PPP_PGINFOSMODIF').AsString;
      If GetControlText('METHODEINIT') = 'RECUP' then RecupHisto(Champ)
      else HistoParDefaut(Champ,True);
      MoveCurProgressForm ('Information récupérée : '+RechDom('PGCHAMPSAL',Champ,False));
    end;
  end;
  FiniMoveProgressForm;
  Liste.ClearSelected;
end;

procedure TOF_PGMULINITHISTO.RecupHisto(Champ : String);
var TobSal,TobAncH,TobNewHisto,TobMaj : Tob;
    TypeHisto,AncValeur,LaValeur,LeType,Tablette,Suffixe : String;
    AncSal,Salarie : String;
    LaDate : TDateTime;
    Q : TQuery;
    i : Integer;
begin
  ExecuteSQL('DELETE FROM PGHISTODETAIL WHERE PHD_PGINFOSMODIF="'+Champ+'"');
  Suffixe := Copy(Champ,5,Length(Champ)-1);
  If Not ExisteSQL('SELECT * FROM DECHAMPS WHERE DH_NOMCHAMP="PHS_'+Suffixe+'"') then
  begin
    HistoParDefaut(Champ,True);
    Exit;
  end;
  LeType := '';
  Tablette := '';
  Q := OpenSQL('SELECT * FROM PARAMSALARIE WHERE PPP_PREDEFINI="CEG" AND PPP_PGINFOSMODIF="'+Champ+'"',True);
  If Not Q.Eof then
  begin
    Tablette := Q.FindField('PPP_LIENASSOC').AsString;
    leType := Q.FindField('PPP_PGTYPEDONNE').AsString;
  end;
  Ferme(Q);
  Q := OpenSQL('SELECT PHS_SALARIE,PHS_COMMENTAIRE,PHS_DATEEVENEMENT,PHS_'+Suffixe+' FROM HISTOSALARIE WHERE PHS_B'+Suffixe+'="X" ORDER BY PHS_SALARIE,PHS_DATEEVENEMENT',True);
  TobAncH := Tob.Create('AncHistorique',Nil,-1);
  TobAncH.LoadDetailDB('AncHistorique','','',Q,False);
  Ferme(Q);
  AncSal := '';
  AncValeur := '';
  TobMaj := Tob.Create('Maj',Nil,-1);
  For i := 0 To TobAncH.Detail.Count - 1 do
  begin
      If Copy(TobAncH.Detail[i].GetValue('PHS_COMMENTAIRE'),1,4) = 'Init' then TypeHisto := '001'
      else TypeHisto := '002';
      Salarie := TobAncH.Detail[i].GetValue('PHS_SALARIE');
      If AncSal <> Salarie then AncValeur := '';
      LaValeur := TobAncH.Detail[i].GetString('PHS_'+Suffixe);
      LaDate := TobAncH.Detail[i].GetValue('PHS_DATEEVENEMENT');
      TobNewHisto := Tob.Create('PGHISTODETAIL',TobMaj,-1);
      TobNewHisto.PutValue('PHD_SALARIE',Salarie);
      TobNewHisto.PutValue('PHD_ETABLISSEMENT','');
      TobNewHisto.PutValue('PHD_ORDRE',-1);
      TobNewHisto.PutValue('PHD_GUIDHISTO',AglGetGuid());
      TobNewHisto.PutValue('PHD_PGINFOSMODIF',Champ);
      TobNewHisto.PutValue('PHD_PGTYPEHISTO',TypeHisto);
      TobNewHisto.PutValue('PHD_ANCVALEUR',AncValeur);
      TobNewHisto.PutValue('PHD_NEWVALEUR',LaValeur);
      TobNewHisto.PutValue('PHD_PGTYPEINFOLS','SAL');
      TobNewHisto.PutValue('PHD_TYPEVALEUR',LeType);
      TobNewHisto.PutValue('PHD_TABLETTE',Tablette);
      TobNewHisto.PutValue('PHD_DATEAPPLIC',LaDate);
      TobNewHisto.PutValue('PHD_TRAITEMENTOK','X');
      TobNewHisto.PutValue('PHD_DATEFINVALID',IDate1900);
//      TobNewHisto.InsertDB(Nil);
//      TobNewHisto.Free;
      AncValeur := LaValeur;
      AncSAL := Salarie;
  end;
  TobMaj.SetAllModifie(True);
  TobMaj.InsertDB(Nil);
  TobMaj.Free;
  TobAncH.Free;
  HistoParDefaut(Champ,False);
end;

procedure TOF_PGMULINITHISTO.HistoParDefaut(Champ : String;PourTous : Boolean);
var TobSal,TobNewHisto : Tob;
    TypeHisto,LaValeur,LeType,Tablette,Suffixe : String;
    Salarie : String;
    LaDate : TDateTime;
    Q : TQuery;
    i : Integer;
    StWhere,Theme,StSelect : String;
begin
  Suffixe := Copy(Champ,5,Length(Champ)-1);
  If suffixe = 'SITUATIONFAMI' then suffixe := 'SITUATIONFAMIL';
  if Not PourTous then StWhere := 'WHERE PSA_SALARIE NOT IN (SELECT PHS_SALARIE FROM HISTOSALARIE WHERE PHS_B'+Suffixe+'="X") '
  Else StWhere := '';
  LeType := '';
  Tablette := '';
  Q := OpenSQL('SELECT * FROM PARAMSALARIE WHERE PPP_PREDEFINI="CEG" AND PPP_PGINFOSMODIF="'+Champ+'"',True);
  If Not Q.Eof then
  begin
    Tablette := Q.FindField('PPP_LIENASSOC').AsString;
    leType := Q.FindField('PPP_PGTYPEDONNE').AsString;
    Theme := Q.FindField('PPP_PGTHEMESALARIE').AsString;
  end;
  Ferme(Q);
  If Suffixe='TAUXPARTIEL' then StSelect := ',PSA_CONDEMPLOI'
  else StSelect := '';
  Q := OpenSQL('SELECT PSA_SALARIE,PSA_DATEENTREE,PSA_'+Suffixe+StSelect+' FROM SALARIES '+StWhere+'ORDER BY PSA_SALARIE',True);
  TobSal := Tob.Create('AncHistorique',Nil,-1);
  TobSal.LoadDetailDB('AncHistorique','','',Q,False);
  Ferme(Q);
  For i := 0 To TobSal.Detail.Count - 1 do
  begin
      LaValeur := TobSal.Detail[i].GetString('PSA_'+Suffixe);
      If Suffixe = 'TAUXPARTIEL' then
      begin
        If TobSal.Detail[i].GetString('PSA_CONDEMPLOI')<>'P' then Continue;
      end;
      If (Theme <> 'PRO') or ((Theme='PRO') and (LaValeur <> '')) then
      begin
        TypeHisto := '001';
        Salarie := TobSal.Detail[i].GetValue('PSA_SALARIE');
        LaDate := TobSal.Detail[i].GetValue('PSA_DATEENTREE');
        TobNewHisto := Tob.Create('PGHISTODETAIL',Nil,-1);
        TobNewHisto.PutValue('PHD_SALARIE',Salarie);
        TobNewHisto.PutValue('PHD_ETABLISSEMENT','');
        TobNewHisto.PutValue('PHD_ORDRE',-1);
        TobNewHisto.PutValue('PHD_GUIDHISTO',AglGetGuid());
        TobNewHisto.PutValue('PHD_PGINFOSMODIF',Champ);
        TobNewHisto.PutValue('PHD_PGTYPEHISTO',TypeHisto);
        TobNewHisto.PutValue('PHD_ANCVALEUR','');
        TobNewHisto.PutValue('PHD_NEWVALEUR',LaValeur);
        TobNewHisto.PutValue('PHD_PGTYPEINFOLS','SAL');
        TobNewHisto.PutValue('PHD_TYPEVALEUR',LeType);
        TobNewHisto.PutValue('PHD_TABLETTE',Tablette);
        TobNewHisto.PutValue('PHD_DATEAPPLIC',LaDate);
        TobNewHisto.PutValue('PHD_TRAITEMENTOK','X');
        TobNewHisto.PutValue('PHD_DATEFINVALID',IDate1900);
        TobNewHisto.InsertDB(Nil);
        TobNewHisto.Free;
      end;
  end;
  TobSal.Free;
end;


{TOF_PGCONFIRMPAIE}

procedure TOF_PGCONFIRMPAIE.OnClose;
begin
  Inherited;
  TFVierge(Ecran).Retour := GetcontrolText('FOUI');
end;

procedure TOF_PGCONFIRMPAIE.OnArgument (S : String ) ;
begin
  Inherited;
  SetControlCaption('LIBELLE1',ReadTokenPipe(S,';'));
  SetControlCaption('LIBELLE2',ReadTokenPipe(S,';'));
end;


Initialization
  registerclasses ( [ TOF_PGMULINITHISTO,TOF_PGCONFIRMPAIE ] ) ;
end.


