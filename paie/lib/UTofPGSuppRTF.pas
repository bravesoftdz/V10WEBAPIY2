{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 10/11/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGEXPORTRTF ()
Mots clefs ... : TOF;PGEXPORTRTF
*****************************************************************}
{
  PT1 11/07/2007  MF  V_72 FQ 14529 Maj ENVOISOCIAL qd suppression fichier DUC
}
Unit UTofPGSuppRTF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
     HDB,
    {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,

{$else}
     eMul,

{$ENDIF}
     entPaie, //PT1
     forms,
     UTOB,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UYFileSTD,
     PGOutils,
     AGLUtilOle,
     HTB97,
     HQry,
     ed_tools,
     P5Util,
     UTOF ; 

Type
  TOF_PGSUPPRTF = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    private
    {$IFNDEF EAGLCLIENT}
    Grille : THDBGrid;
    {$ELSE}
    Grille : THGrid;
    {$ENDIF}
    QMul : THQuery;
    procedure SuppFichiers(Sender : TObject);
    procedure SuppressionDonnees(NomFic,Langue,Predef,NoDossier : String);
  end ;

Implementation

procedure TOF_PGSUPPRTF.OnArgument (S : String ) ;
var BSupp : TToolBarButton97;
begin
  Inherited ;
  {$IFNDEF EAGLCLIENT}
        Grille := THDBGrid (GetControl ('Fliste'));
        {$ELSE}
        Grille := THGrid (GetControl ('Fliste'));
        {$ENDIF}
        QMul := THQuery(Ecran.FindComponent('Q'));
        BSupp := TToolBarButton97(GetControl('BSUPP'));
        If BSupp <> Nil then BSupp.OnClick := SuppFichiers;
        SetControlVisible('YFS_CRIT3',False);
        SetControlVisible('TYFS_CRIT3',False);
        If V_PGI.ModePcl = '1' then
        begin
             If Not DroitCegPaie then SetControlText('XX_WHERE','YFS_PREDEFINI<>"CEG" AND NOT (YFS_PREDEFINI="DOS" AND YFS_NODOSSIER<>"'+V_PGI.NoDossier+'")')
             else SetControlText('XX_WHERE','NOT (YFS_PREDEFINI="DOS" AND YFS_NODOSSIER<>"'+V_PGI.NoDossier+'")');
        end
        else If Not DroitCegPaie then SetControlText('XX_WHERE','YFS_PREDEFINI<>"CEG"');
end ;

procedure TOF_PGSUPPRTF.SuppFichiers(Sender : TObject);
var NomF,Langue,Predef,NoDossier : String;
    i : Integer;
begin
  if (Grille.NbSelected = 0) and (TFMul(Ecran).BSelectAll.Down = False) then
  begin
    PGIBox('Aucun élément sélectionné', Ecran.Caption);
    Exit;
  end;
  if ((Grille.nbSelected) > 0) and (not Grille.AllSelected) then
  begin
       for i := 0 to Grille.NbSelected - 1 do
       begin
            Grille.GotoLeBookmark(i);
            {$IFDEF EAGLCLIENT}
            TFMul(Ecran).Q.TQ.Seek(Grille.Row - 1);
            {$ENDIF}
            NomF := QMul.FindField('YFS_NOM').AsString;
            Langue := QMul.FindField('YFS_LANGUE').AsString;
            Predef := QMul.FindField('YFS_PREDEFINI').AsString;
            NoDossier := QMul.FindField('YFS_NODOSSIER').AsString;
            SuppressionDonnees(NomF,Langue,Predef,NoDossier);
       end;
       Grille.ClearSelected;
  end;
  if (Grille.AllSelected = TRUE) then
  begin
     {$IFDEF EAGLCLIENT}
     if (TFMul(Ecran).bSelectAll.Down) then
     TFMul(Ecran).Fetchlestous;
     {$ENDIF}
     QMul.First;
     while not QMul.EOF do
     begin
          NomF := QMul.FindField('YFS_NOM').AsString;
          NomF := QMul.FindField('YFS_NOM').AsString;
          Langue := QMul.FindField('YFS_LANGUE').AsString;
          Predef := QMul.FindField('YFS_PREDEFINI').AsString;
          NoDossier := QMul.FindField('YFS_NODOSSIER').AsString;
          SuppressionDonnees(NomF,Langue,Predef,NoDossier);
          QMul.Next;
     end;
     Grille.AllSelected := False;
  end;
  TFMul(Ecran).BCherche.Click;
end;

procedure TOF_PGSUPPRTF.SuppressionDonnees(NomFic,Langue,Predef,NoDossier : String);
var Q : TQuery;
    Guid, Crit1 : String;     //PT1
    FicSuppr : string; //PT1
begin
     Guid := '';
{* d PT1
     Q := OpenSQL('SELECT YFS_FILEGUID FROM YFILESTD WHERE YFS_CODEPRODUIT="PAIE" AND '+
     'YFS_NOM="'+NomFic+'" AND YFS_LANGUE="'+Langue+'" AND YFS_PREDEFINI="'+Predef+'" AND YFS_NODOSSIER="'+NoDossier+'"',True);
     If Not Q.Eof then Guid := Q.FindField('YFS_FILEGUID').AsString;*}

     Q := OpenSQL('SELECT YFS_FILEGUID, YFS_CRIT1 FROM YFILESTD WHERE YFS_CODEPRODUIT="PAIE" AND '+
     'YFS_NOM="'+NomFic+'" AND YFS_LANGUE="'+Langue+'" AND YFS_PREDEFINI="'+Predef+'" AND YFS_NODOSSIER="'+NoDossier+'"',True);
     If Not Q.Eof then
     begin
        Guid := Q.FindField('YFS_FILEGUID').AsString;
        Crit1 := Q.FindField('YFS_CRIT1').AsString;
     end;
// f PT1
     ferme(Q);
     ExecuteSQL('DELETE FROM YFILESTD WHERE YFS_FILEGUID="'+Guid+'"');
     ExecuteSQL('DELETE FROM NFILES WHERE NFI_FILEGUID="'+Guid+'"');
     ExecuteSQL('DELETE FROM NFILEPARTS WHERE NFS_FILEGUID="'+Guid+'"');
// d PT1
     if (Crit1 = 'DUC') or (Crit1 ='DAD') then
     // mise à jour de la table ENVOISOCIAL
     begin
       try
        begintrans;
        if (Crit1 = 'DUC') then
        // dans le cas DUCS EDI : possibilité de supprimer le fichier .cop associé
        begin
          Q := OpenSQL('SELECT PES_FICHIEREMIS FROM ENVOISOCIAL '+
                       'WHERE PES_GUID1 = "'+Guid+'"',True);
          If Not Q.Eof then
            if (Q.FindField('PES_FICHIEREMIS').AsString  <> '') then
               if (PGIAsk ('Mise à jour de la table des envois'+ #13#10+
                           ' pour la déclaration '+ NomFic+#13#10+'.'+
                           'Voulez-vous supprimer le fichier envoyé? '+
                           Q.FindField('PES_FICHIEREMIS').AsString,
                           Ecran.Caption) = mrYes) then
               begin
                 FicSuppr := VH_Paie.PGCheminEagl+'\'+Q.FindField('PES_FICHIEREMIS').AsString;
                 if FileExists(FicSuppr) then
                 begin
                   DeleteFile(PChar(FicSuppr));
                 end;
               end;

          ferme(Q);
        end;

        ExecuteSQL('DELETE FROM ENVOISOCIAL WHERE PES_GUID1 = "'+Guid+'"');

        CommitTrans;
       except
        Rollback;
       end;
     end;
// f PT1
end;

Initialization
  registerclasses ( [ TOF_PGSUPPRTF ] ) ;
end.

