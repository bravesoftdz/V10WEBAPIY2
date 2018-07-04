{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 05/12/2001
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : MBOPANIER_MOYEN ()
Mots clefs ... : TOF;MBOPANIER_MOYEN
*****************************************************************}
Unit UTOFMBOPANIER_MOYEN ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     eQRS1,
{$ELSE}
     db,
     dbtables,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     UTOB,
     HQry,
     QRS1,
     HStatus,
     M3FP ;

Type
  TOF_MBOPANIER_MOYEN = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure OnChange (Sender : TObject) ;
  private
    ComboModeleEtat : TComboBox ;
    stWhere1        : string ;
  end ;

Implementation

procedure TOF_MBOPANIER_MOYEN.OnChange (Sender : TObject) ;
begin
  if TFQRS1(Ecran).CodeEtat='ET1' then
     BEGIN
     SetControlVisible('CBOX1', True)  ;
     SetControlVisible('CBOX2', True)  ;
     SetControlVisible('CBOX3', True) ;
     END else
  if TFQRS1(Ecran).CodeEtat='ET2' then
     BEGIN
     SetControlVisible('CBOX1', False) ;
     SetControlVisible('CBOX2', False) ;
     SetControlVisible('CBOX3', True)  ;
     END ;
end ;

procedure TOF_MBOPANIER_MOYEN.OnUpdate ;
var i1, iHeure : integer ;
    stSQL, stHeure, stJour : String ;
    TobTemp, TobPM, TobFillePM : TOB ;
    QPM : TQuery ;
begin
  Inherited ;
  // Initialisations
  stWhere1 := StringReplace(stWhere1,'GZP_','GP_',[rfReplaceAll]) ;

  ExecuteSQL('DELETE FROM GCTMPPANIERMOYEN WHERE GZP_UTILISATEUR = "'+V_PGI.USer+'"');

  stSQL := 'SELECT GL_ETABLISSEMENT, GL_REPRESENTANT, GL_NUMERO, GL_QTEFACT, ' +
                  'GL_TOTALTTCDEV, GL_TOTREMLIGNEDEV, GL_DATEPIECE, GL_PMAP, ' +
                  'GL_DATEMODIF, GL_TOTALHTDEV ' +
           'FROM PIECE LEFT JOIN LIGNE '+
                 'ON GL_NATUREPIECEG=GP_NATUREPIECEG ' +
                 'AND GL_SOUCHE=GP_SOUCHE ' +
                 'AND GL_NUMERO=GP_NUMERO ' +
                 'AND GL_INDICEG=GP_INDICEG ' + 
                 'WHERE GP_NATUREPIECEG="FFO" AND ' + copy(stwhere1,7,length(stwhere1)-7) +
                 ' AND GL_TYPEARTICLE="MAR" ' ;
  QPM := OpenSQL(stSQL, True) ;

  TObTemp := Nil ; TobPM := Nil ;
  if not QPM.EOF then
    BEGIN
    TobTemp := TOB.Create('', Nil, -1) ;
    TobTemp.LoadDetailDB('', '', '', QPM, False) ;
    TobPM := Tob.Create('PANIER MOYEN', Nil, -1) ;
    InitMove(TobTemp.Detail.Count,'') ;
    For i1 := 0 to TobTemp.Detail.Count - 1 do
      BEGIN
        TobFillePM := TOB.Create('GCTMPPANIERMOYEN', TobPM, -1) ;
        TobFillePM.PutValue('GZP_UTILISATEUR', V_PGI.User) ;
        TobFillePM.PutValue('GZP_COMPTEUR', i1) ;
        TobFillePM.PutValue('GZP_ETABLISSEMENT', TobTemp.Detail[i1].GetValue('GL_ETABLISSEMENT')) ;
        TobFillePM.PutValue('GZP_VENDEUR', TobTemp.Detail[i1].GetValue('GL_REPRESENTANT')) ;
        TobFillePM.PutValue('GZP_NUMERO', TobTemp.Detail[i1].GetValue('GL_NUMERO')) ;
        TobFillePM.PutValue('GZP_QTEFACT', TobTemp.Detail[i1].GetValue('GL_QTEFACT')) ;
        TobFillePM.PutValue('GZP_TOTALTTCDEV', TobTemp.Detail[i1].GetValue('GL_TOTALTTCDEV')) ;
        TobFillePM.PutValue('GZP_TOTREMLIGNEDEV', TobTemp.Detail[i1].GetValue('GL_TOTREMLIGNEDEV')) ;
        TobFillePM.PutValue('GZP_PMAP', TobTemp.Detail[i1].GetValue('GL_PMAP')) ;
        TobFillePM.PutValue('GZP_TOTALHTDEV', TobTemp.Detail[i1].GetValue('GL_TOTALHTDEV')) ;
        TobFillePM.PutValue('GZP_DATEPIECE', TobTemp.Detail[i1].GetValue('GL_DATEPIECE')) ;
        TobFillePM.PutValue('GZP_DATEMODIF', TobTemp.Detail[i1].GetValue('GL_DATEMODIF')) ;
        stHeure := FormatDateTime('dddd;h', TobTemp.Detail[i1].GetValue('GL_DATEPIECE')) ;
        stJour := ReadTokenSt(stHeure) ;
        iHeure := StrToInt(stHeure) ;
        if iHeure < 10 then  iHeure := 9 else  // ventes avant 10 heures
        if iHeure > 21 then  iHeure := 22 ;    // ventes après 21 heures
        TobFillePM.PutValue('GZP_JOUR', stJour) ;
        TobFillePM.PutValue('GZP_TRANCHEHORAIRE', iHeure) ;
        MoveCur(False)
      END ;
    END ;
  FiniMove ;
  Ferme(QPM) ;

  if TobPM <> Nil then TobPM.InsertDB(Nil, True) ;
  TobPM.Free ;
  TobTemp.Free ;

end ;

procedure TOF_MBOPANIER_MOYEN.OnArgument (S : String ) ;
begin
  Inherited ;
  TFQRS1(Ecran).CodeEtat := 'ET1' ;
  TFQRS1(Ecran).FEtat.Text := 'ET1' ;
  ComboModeleEtat := TComboBox(GetControl('FEtat'));
  ComboModeleEtat.OnChange := OnChange ;
  if TFQRS1(Ecran).CodeEtat='ET1' then
     BEGIN
     SetControlVisible('CBOX1', True)  ;
     SetControlVisible('CBOX2', True)  ;
     SetControlVisible('CBOX3', True) ;
     END else
  if TFQRS1(Ecran).CodeEtat='ET2' then
     BEGIN
     SetControlVisible('CBOX1', False) ;
     SetControlVisible('CBOX2', False) ;
     SetControlVisible('CBOX3', True)  ;
     END ;
  SetActiveTabSheet('Standards') ;
end ;

procedure TOF_MBOPANIER_MOYEN.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_MBOPANIER_MOYEN.OnLoad ;
begin
  Inherited ;
  SetControlText('XX_WHERE','') ;
  stWhere1 := RecupWhereCritere(TPageControl(TFQRS1(Ecran).FindComponent('PAGES')));
  SetControlText('XX_WHERE','GZP_UTILISATEUR="'+V_PGI.USer+'"');
end ;

Initialization
  registerclasses ( [ TOF_MBOPANIER_MOYEN ] ) ;
end.
