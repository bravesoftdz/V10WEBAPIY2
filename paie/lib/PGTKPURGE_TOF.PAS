{***********UNITE*************************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 30/06/2003
Modifié le ... : 01/07/2003
Description .. : Source TOF de la FICHE : TICKETPURGE ()
Suite ........ : Purge des commandes (table cdetickets) selon critères 
Suite ........ : (période et/ou date de commande)
Suite ........ : Cas code client = code rattachement : purge du récapitulatif 
Suite ........ : (table RECAPCDE TICKET) + Suppression des fichiers de 
Suite ........ : commandes correspondant.
Mots clefs ... : TOF;PGTKPURGE
*****************************************************************}
{   PT1    MF    14/02/2005  V_6.0 Le nom du pré-fichier de commande contient
                                   les dates de prériode (--> autant de pré-fichiers
                                   que de périodes)
}
Unit PGTKPURGE_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
{$ENDIF}
     EntPaie,
//unused     FileCtrl,
//unused     forms,
     sysutils,
     ComCtrls,
     HTB97,
     HCtrls,
     HEnt1,
     HMsgBox,
     ParamSoc,
     UTOB,
     UTOF ;

Type
  TOF_PGTKPURGE = Class (TOF)
    private
    DateDeb,DateFin,DateCde            : THEdit;
    BtnLance                           : TToolbarButton97;
    Trace,TraceErr                     : TListBox;
    procedure PurgeCommande(Sender: TObject);

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

procedure TOF_PGTKPURGE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGTKPURGE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGTKPURGE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGTKPURGE.OnLoad ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 01/07/2003
Modifié le ... :   /  /    
Description .. : procédure OnArgument 
Mots clefs ... : PAIE ; PGTKPURGE
*****************************************************************}
procedure TOF_PGTKPURGE.OnArgument (S : String ) ;
begin
  Inherited ;
  DateDeb := ThEdit(getcontrol('DATEDEB'));
  DateFin := ThEdit(getcontrol('DATEFIN'));
  DateCde := ThEdit(getcontrol('DATECDE'));

  BtnLance:=TToolbarButton97 (GetControl ('BLANCE'));
  if BtnLance<>NIL then
    BtnLance.OnClick := PurgeCommande;
end ;

procedure TOF_PGTKPURGE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PGTKPURGE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGTKPURGE.OnCancel () ;
begin
  Inherited ;
end ;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 01/07/2003
Modifié le ... :   /  /    
Description .. : procédure PurgeCommande
Suite ........ : Contrôle que les lignes de commandes concernées ont bien 
Suite ........ : été envoyées et intégrées. Dans ces 2 cas affichage d'un 
Suite ........ : message non bloquant.
Suite ........ : Suppression de lignes de CDETICKETS, 
Suite ........ : RECAPCDETICKET
Suite ........ : Suppression physique des fichiers correspondant
Mots clefs ... : PAIE ; PGTKPURGE
*****************************************************************}
procedure TOF_PGTKPURGE.PurgeCommande(Sender: TObject);
var
  TobCde,TobCdeFille                    : TOB;
  Pan                                   : TPageControl;
  Tbsht                                 : TTabSheet;
  Integr, Cde, CdeExist                 : boolean;
  rep, NbLignes                         : Integer;
  St, StWhere, NomFich                  : string;
  Q                                     : TQuery;
// d PT1-1
  Dir                                   : string;
  Periode                               : string;
  sr                                    : TsearchRec;
  FileAttrs, ret                        : integer;
// f PT1-1 
begin
  Integr := False;
  Cde := False;
  CdeExist := False;
  NbLignes := 0;
  St := '';
  StWhere := '';

  Pan := TPageControl (GetControl ('PANELPURGE'));
  Tbsht := TTabSheet (GetControl ('TBSHTTRACE'));
  Trace := TListBox (GetControl ('LSTBXTRACE'));
  TraceErr := TListBox (GetControl ('LSTBXERROR'));
  Trace.Clear;
  TraceErr.Clear;

  if (Trace = NIL) or (TraceErr = NIL) then
  begin
    PGIBox ('La purge de commande ne peut pas être lancée', 'Les composants trace ne sont pas disponibles');
    exit;
  end;
  TobCde := TOB.Create ('La commande',NIL,-1);
  St := 'SELECT * '+
        'FROM CDETICKETS '+
        'WHERE ';
  if (Datedeb.text <> '  /  /    ') and (Datefin.text <> '  /  /    ') then
//PT1-1   StWhere := 'PRT_DATEDEBUT = "'+USDateTime(StrToDate(Datedeb.text))+'" '+
//PT1-1          'AND PRT_DATEFIN = "'+USDateTime(StrToDate(Datefin.text))+'"';
    StWhere := 'PRT_DATEDEBUT >= "'+USDateTime(StrToDate(Datedeb.text))+'" '+
          'AND PRT_DATEFIN <= "'+USDateTime(StrToDate(Datefin.text))+'"';
  if (DateCde.text <> '  /  /    ') then
  begin
    if (Datedeb.text <> '  /  /    ') and (Datefin.text <> '  /  /    ') then
      StWhere := StWhere + ' AND ';
    StWhere := StWhere + 'PRT_DATECDE = "'+USDateTime(StrToDate(Datecde.text))+'"';
  end;
  St := St + StWhere ;
  if (StWhere <> '') then
  begin
    Q := OpenSql (St, TRUE);
    if not Q.EOF then
    begin
      TobCde.LoadDetailDB ('COMMANDE TICKETS', '','',Q,FALSE,FALSE);
      CdeExist := True;
    end
    else
    begin
      PGIBox ('Aucune ligne de commande pour cette période ', 'Pas de purge');
      TraceErr.Items.Add ('Aucune ligne de commande pour cette période');
      Tbsht := TTabSheet (GetControl ('TBSHTERROR'));
    end;
    Ferme(Q);
  end;

  if (CdeExist) then
  begin
    TobCdeFille := TobCde.FindFirst ([''],[''], True);
    while (TobCdeFille  <> Nil) do
    begin
      NbLignes := NbLignes + 1;
      if (TobCdeFille.GetValue('PRT_DATEINTEG') <> IDate1900) then
        Integr := True;
      if (TobCdeFille.GetValue('PRT_DATECDE') <> IDate1900) then
        Cde := True;
      TobCdeFille := TobCde.FindNext ([''],[''], True);
    end;

    if (Integr = False) then
    begin
      rep:=PGIAsk ('Commande non intégrée dans la paie, validez-vous la purge?', Ecran.Caption) ;
      if rep=mrNo then exit ;
      if rep=mrCancel then exit;
    end;
    if (Cde = False) then
    begin
      rep:=PGIAsk ('Les lignes de la commande n''ont pas fait l''objet d''un envoi, validez-vous la purge?', Ecran.Caption) ;
      if rep=mrNo then exit ;
      if rep=mrCancel then exit;
    end;

    try
      BeginTrans;
      St := 'DELETE FROM CDETICKETS WHERE ';
      St := St + StWhere ;
      ExecuteSQL(St);
      Trace.Items.Add ('Purge de '+IntToStr(NbLignes)+
                       ' lignes de commandes pour la période du '+
                       DateDeb.text+
                       ' au '+DateFin.text);

//d PT1-1
    Dir :=  Trim(GetParamSocSecur('SO_PGREPERTTICKET',''));

    Periode := Copy(Datedeb.text,9,2)+Copy(Datedeb.text,4,2)+Copy(Datedeb.text,1,2)+
               Copy(Datefin.text,9,2)+Copy(Datefin.text,4,2)+Copy(Datefin.text,1,2) ;

    FileAttrs := 0;
    FileAttrs := FileAttrs + faAnyFile;
    ret := FindFirst(Dir+'\P*'+Periode+'.st1',FileAttrs,sr);
    if (ret = 0) then
      while (ret = 0) do
      begin
        if (FileExists(Dir+'\'+ sr.Name)) then
         DeleteFile(PChar(Dir+'\'+ sr.Name));
        ret := FindNext(sr);
      end;  {fin while (ret = 0)}
    sysutils.FindClose(sr);

//f PT1-1

      CommitTrans;
    except
      Rollback;
      PGIBox ('Une erreur est survenue lors de la purge de commande','');
      TraceErr.Items.add ('Une erreur est survenue lors de la purge de commande');
      Tbsht := TTabSheet (GetControl ('TBSHTERROR'));
      CdeExist := False;
    end;
  end;

  rep := mrNo;
  StWhere := ''; // PT1-1
  {Suppression physique des fichiers de commande }
  St := 'SELECT PHT_NOMFICH FROM RECAPCDETICKET WHERE ';
//PT1-1 StWhere := 'PHT_DATEDEBUT = "'+USDateTime(StrToDate(Datedeb.text))+'" '+
//             'AND PHT_DATEFIN = "'+USDateTime(StrToDate(Datefin.text))+'"';
  if (DateCde.text <> '  /  /    ') then
  begin
    if (Datedeb.text <> '  /  /    ') and (Datefin.text <> '  /  /    ') then
      StWhere := 'PHT_DATEDEBUT = "'+USDateTime(StrToDate(Datedeb.text))+'" '+  //PT1-1
             'AND PHT_DATEFIN = "'+USDateTime(StrToDate(Datefin.text))+'" AND '; //PT1-1

//PT1-1      StWhere := StWhere + ' AND ';
    StWhere := StWhere + 'PHT_DATECDE = "'+USDateTime(StrToDate(Datecde.text))+'"';
  end;
  if (StWhere <> '') then
  begin
    St := St + StWhere;
    Q := OpenSql (St, TRUE);
    if not Q.EOF then
    begin
      if (VH_Paie.PGCodeClient = VH_Paie.PGCodeRattach) and (CdeExist) and (Cde) then
      begin
        rep:=PGIAsk ('Voulez-vous purger les lignes associées du récapitulatif?', Ecran.Caption) ;
      end
      else
        if (VH_Paie.PGCodeClient = VH_Paie.PGCodeRattach) and (not CdeExist) then
        begin
          rep:=PGIAsk ('Il existe des lignes dans le récapitulatif répondant à ces '+
                       'critères. Voulez-vous les purger?', Ecran.Caption) ;
        end;

      if rep=mrYes then
      begin
        Q.First ;
        while not Q.Eof do
        begin
          NomFich := Q.FindField ('PHT_NOMFICH').Asstring;
          if (FileExists(NomFich)) then
            DeleteFile(PChar(NomFich));
          Q.Next ;
        end;
      end;
    end;
    Ferme(Q);
  end;

  if rep=mrYes then
  begin
   {Purge de la table RECAPCDETICKET}
   try
     BeginTrans ;
     St := 'DELETE FROM RECAPCDETICKET WHERE ';
     if (StWhere <> '') then
     begin
       St := St + StWhere ;
       ExecuteSQL(St);
     end;
     Trace.Items.add ('La purge du récapitulatif a été éffectuée');
     CommitTrans;
   except
     Rollback;
     PGIBox ('Une erreur est survenue lors de la purge du récapitulatif','');
     TraceErr.Items.add ('Une erreur est survenue lors de la purge du récapitulatif');
     Tbsht := TTabSheet (GetControl ('TBSHTERROR'));
   end;
  end;

  Pan.ActivePage := Tbsht;
  FreeAndNil(TobCde);

end;

Initialization
  registerclasses ( [ TOF_PGTKPURGE ] ) ;
end.
