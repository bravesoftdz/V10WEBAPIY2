{***********UNITE*************************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 20/07/2004
Modifié le ... : 27/07/2004
Description .. : Source TOF de la FICHE : CRITMAINTIEN_MUL ()
Suite ........ : Multi critère de séléction des tables CRITMAINTIEN 
Suite ........ : (méthodes de calcul du maintien)
Mots clefs ... : TOF;CRITMAINTIEN_MUL; MAINTIEN
*****************************************************************}
{
PT1 30/11/2004 MF V_60 Correction traitement maintien qd champ catégorie renseigné.
PT2 23/03/2005 MF V_60 Correction faute orthographe FQ 12112
PT3 20/03/2006 MF V_65 FQ 12971 Qd champ secondaire de séléction desactivé, possibilité
                       d'initialiser le champ PCM_VALCATEGORIE de chaque critère.
PT4 11/05/2006 MF V_65 FQ 12971 : correction UPDATE (Ajout ##PCM_PREDEFINI##)
PT5 13/06/2006 MF V_70 FQ 12971 : correction SELECT (Ajout ##PCM_PREDEFINI##)
PT6 21/02/2007 MF V_702 correction : VALCATEGORIE remplace VALCATEG 
}
Unit CRITMAINTIEN_MUL_TOF ;

Interface

Uses
// unused    StdCtrls, 
     Controls,
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
//unused     HDB,
     mul,
{$else}
     eMul,
     uTob,
{$ENDIF}
     Entpaie,
//unused     forms,
//unused     HTB97,
//unused     sysutils,
//unused     ComCtrls,
     HCtrls,
     Hent1,           // PT3
     HmsgBox,         // PT3
//unused     HEnt1,
//unused     HMsgBox,
//unused     PgOutils2,       
     UTOF ; 

Type
  TOF_CRITMAINTIEN_MUL = Class (TOF)
  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation

procedure TOF_CRITMAINTIEN_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CRITMAINTIEN_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CRITMAINTIEN_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CRITMAINTIEN_MUL.OnLoad ;
var
  Q                                    : TQuery;
  st                                   : string;
  reponse                              : integer;
begin
  Inherited ;

// d PT3
  if (VH_Paie.PGCritMaintien = '') then
  begin
    st := 'SELECT PCM_VALCATEGORIE FROM CRITMAINTIEN WHERE ##PCM_PREDEFINI## PCM_VALCATEGORIE <> ""';     // PT5
    Q := OpenSql(st, TRUE);
    if not Q.eof then
    begin
      reponse:=PGIAsk('Le champ secondaire de sélection est desactivé.#13#10'+
               'Voulez-vous initialiser ce champ pour chaque critère ?',TFMul(Ecran).Caption);
      if (reponse = mrYes) then
      begin
        try
        BeginTrans;
          ExecuteSql('UPDATE CRITMAINTIEN SET PCM_VALCATEGORIE = ""'+
                     ' WHERE ##PCM_PREDEFINI## PCM_VALCATEGORIE<>""');    // PT4
        CommitTrans;
        except
        Rollback;
        PGIBox('Maj annulée : Echec lors de la maj du critère.', 'mise à jour');
        end;
      end;

    end;
    ferme(Q);
  end;
// f PT3
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 27/07/2004
Modifié le ... :   /  /    
Description .. : Procédure OnArgument 
Suite ........ : Chargement de la fiche.
Suite ........ : L'affichage du critère catégorie est fonction du paramètre 
Suite ........ : société PGCritMaintien
Mots clefs ... : PAIE ; MAINTIEN
*****************************************************************}
procedure TOF_CRITMAINTIEN_MUL.OnArgument (S : String ) ;
var
  Q                                             : TQuery;
  Categ                                         : string;
begin
  Inherited ;
// d PT2
   TFMul(Ecran).Caption := 'Critères de maintien du salaire';
   UpdateCaption(TFMul(Ecran));
// f PT2

  if (VH_Paie.PGCritMaintien = '') then
  begin
    SetControlEnabled('PCM_VALCATEGORIE', False);    //PT6
    SetControlVisible('PCM_VALCATEGORIE', False);    //PT6
    SetControlVisible('TPCM_VALCATEGORIE', False);   //PT6
  end
  else
  begin
    Q := OpenSQL('SELECT PAI_LIENASSOC, PAI_LIBELLE FROM PAIEPARIM WHERE PAI_IDENT = "'+ //PT1
               VH_Paie.PGCritMaintien+'"' , TRUE);
    if not Q.EOF then
    begin
      Categ := Q.Fields[0].AsString;
      setControlProperty('PCM_VALCATEGORIE','DataType',Categ);      //PT6
      setControlProperty('TPCM_VALCATEGORIE','Caption',Q.Fields[1].AsString);//PT1  //PT6
    end
    else
      Categ := '';
    Ferme(Q);

  end;
end ;

procedure TOF_CRITMAINTIEN_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CRITMAINTIEN_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CRITMAINTIEN_MUL.OnCancel () ;
begin
  Inherited ;
end ;
Initialization
  registerclasses ( [ TOF_CRITMAINTIEN_MUL ] ) ;
end.
