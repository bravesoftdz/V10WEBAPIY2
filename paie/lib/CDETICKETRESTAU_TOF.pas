{***********UNITE*************************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 21/05/2003
Modifié le ... : 11/06/2003
Description .. : Source TOF de la FICHE : CDETICKETRESTAU ()
Suite ........ : Confection du fichier final de commande de tickets
Suite ........ : restaurant
Mots clefs ... : TOF;CDETICKETRESTAU
*****************************************************************}
{
  PT1    MF    14/02/2005  V_6.0   : 1- La confection du fichier final n'est
                                     possible que si il existe des Cdes dans
                                     CDETICKET pour la période.
                                     2- Mise en place de la confection du fichier
                                     pour NATEXIS
                                     3- Le nom du pré-fichier de commande contient
                                     les dates de période (--> autant de pré-fichiers
                                     que de périodes)
                                     4- Pour NATEXIS édition du document à faxer
  PT2    MF    22/01/2007  V_720     Nouveau fournisseur : ACCOR
  PT3    NA    02/06/2008  V_850     Nouveau fournisseur : Chéque déjeuner
  PT4    NA    23/06/2008  V_850     Si chèque déjeuner et client/etabl : faire un fichier final par code client
}
Unit CDETICKETRESTAU_TOF ;

Interface

Uses
     {$IFDEF VER150}
     Variants,
     {$ENDIF}
     StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
//unused    db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     EdtREtat,
     FE_Main,
{$ELSE}
     MaineAGL,
     UtileAGL,
{$ENDIF}
//unused    forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     UTOB,
     HTB97,
//unused     PgOutils,
     EntPaie,
     ParamSoc,
     FileCtrl,
     Windows,
     PgOutils2,
     Mailol ;

Type
  TOF_CDETICKETRESTAU = Class (TOF)
    private
    DateDeb,DateFin,DateCde            : THEdit;
    BtnLance                           : TToolbarButton97;
    Trace,TraceErr                     : TListBox;
    Fournisseur                        : string; // PT1-2
    Pages                              : TPageControl; // PT1-4
    Supp                               : string;       // PT1-4
    Wtotalticket                       : integer; // pt3

  //pt4  procedure OuvreFichierFinal(var FFichierFinal : TextFile; var FichierFinal : string; Dir : string; var CodeRetour : boolean);
    procedure OuvreFichierFinal(var FFichierFinal : TextFile; var FichierFinal : string; Dir : string; var CodeRetour : boolean; codeclient : string); // pt4
    procedure ConfFichierFinal(codeclient : string); // pt4
   // pt4 procedure ConfFichierFinal(Sender: TObject);
    procedure Lanceprocedure(Sender: TObject);   // pt4
  // pt4  procedure CreatLigneCdeFinale(Dir :string; var FFichierFinal : TextFile; FichierFinal : string;var CodeRetour : boolean);
    procedure CreatLigneCdeFinale(Dir :string; var FFichierFinal : TextFile; FichierFinal : string;var CodeRetour : boolean; codeclient : string);  // pt4

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

procedure TOF_CDETICKETRESTAU.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CDETICKETRESTAU.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CDETICKETRESTAU.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CDETICKETRESTAU.OnLoad ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 11/06/2003
Modifié le ... : 11/06/2003
Description .. : Procédure OnArgument
Suite ........ : La période proposée est la période en cours, elle est 
Suite ........ : modifiable
Suite ........ : Pour info affichage du destinataire, de l'autaeur de la 
Suite ........ : commande (Utilisateur)
Suite ........ : la date de commande proposée est la date du jour (date 
Suite ........ : système), elle est modifiable.
Suite ........ : Le support de transmission est par défaut "télétrans 
Suite ........ : internet",  il est modifiable.
Suite ........ : Le contôle Saisie de stock est par défaut positionné à faux
Suite ........ : formatage du nom de fichier final.
Mots clefs ... : PAIE; CDETICKETRESTAU
*****************************************************************}
procedure TOF_CDETICKETRESTAU.OnArgument (S : String ) ;
var
  DebPer,FinPer,ExerPerEncours              : String;
  Destinataire                              : THEdit;
  OkOk                                      : Boolean;
  begin
  if (VH_Paie.PgCodeRattach = '') or (VH_Paie.PgCodeRattach = NULL) then
  begin
    PgiBox('le code de rattachement n''est pas renseigné',
           'Commande de titres restaurant');
    exit;
  end;

  Inherited ;
  Pages := TPageControl(GetControl('Pages')); // PT1-4
  Fournisseur := GetParamSocSecur('SO_PGTYPECDETICKET',''); // PT1-2
  DateDeb := ThEdit(getcontrol('DATEDEBUT'));
  DateFin := ThEdit(getcontrol('DATEFIN'));
  DateCde := ThEdit(getcontrol('DATECDE'));
  OkOk := RendPeriodeEnCours (ExerPerEncours,DebPer,FinPer);
  if OkOk then
  begin
   if  DateDeb <> NIL   then DateDeb.text:= DebPer;
   if  DateFin <> NIL   then DateFin.text:= FinPer;
  end;

  if DateCde <> NIl then DateCde.text := DateToStr(Date);

  Destinataire :=  ThEdit(getcontrol('DESTINATAIRE'));
  if (Destinataire <> NIl) then
    Destinataire.text := RechDom('PGTYPETICKET',
                                 GetParamSocSecur('SO_PGTYPECDETICKET',''),
                                 FALSE);

  SetControlText('SUPPORTEMIS','TEL');
  SetControlText ('AUTEUR',V_PGI.UserName);

  BtnLance:=TToolbarButton97 (GetControl ('BLANCE'));
  if BtnLance<>NIL then
  // pt4  BtnLance.OnClick := ConfFichierFinal;
  BtnLance.OnClick := Lanceprocedure;

// d PT1-2
  if (Fournisseur = '001') then
  begin
  // SODEXHO
    SetControlText ('NOMFIC',
                    copy('00000000',1,8-length(Trim(VH_Paie.PgCodeRattach)))+    // PT2
                    VH_Paie.PgCodeRattach+
                    '.st1');
  end
  else
  if (Fournisseur = '002') then           // PT2
  begin
  // NATEXIS
    SetControlText ('NOMFIC',
                    'E'+
                    copy('0000000',1,7-length(Trim(VH_Paie.PgCodeRattach)))+
                    VH_Paie.PgCodeRattach+
                    '.TXT');
// d PT2
  end
  else
  if (Fournisseur = '003') then
  begin
  // ACCOR
    SetControlText ('NOMFIC',
                     copy('00000000',1,6-length(Trim(VH_Paie.PgCodeRattach)))+
                    VH_Paie.PgCodeRattach+
                    '.csv');
// f PT2
// pt3  end;
// deb pt3
  end
  else
  if (Fournisseur = '004') then
  begin
  // Chèque déjeuner
    if (Vh_Paie.PGFACTETABL) then
    begin
     // code client / etabl : rendre invisible le nom du fichier et la saisie des stocks
      setcontrolvisible('NOMFIC', false);
      setcontrolenabled('NOMFIC', false);
      setcontrolvisible('LNOMFIC', false);
      setcontrolvisible('STOCK', false);
    end
    else
   
     SetControlText ('NOMFIC',
                     copy('00000',1,5-length(Trim(VH_Paie.PgCodeRattach)))+
                    VH_Paie.PgCodeRattach+
                    '.prn');
  end;
  // fin pt7
// f PT1-2
  SetControlChecked ('STOCK',FALSE);
end ;

procedure TOF_CDETICKETRESTAU.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CDETICKETRESTAU.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CDETICKETRESTAU.OnCancel () ;
begin
  Inherited ;
end ;

// deb pt4
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - NA
Créé le ...... : 24/06/2008
Modifié le ... : 24/06/2008
Description .. : procédure Lanceprocedure
Suite ........ : Lance la procédure de confection du (ou des) fichiers de commande
Mots clefs ... : PAIE; CDETICKETRESTAU
*****************************************************************}
procedure TOF_CDETICKETRESTAU.Lanceprocedure(Sender: TObject);
var
Q : TQuery;
codeclient : string;

begin
  Trace := TListBox (GetControl ('LSTBXTRACE'));
  TraceErr := TListBox (GetControl ('LSTBXERROR'));
//  Trace.Clear;
  Trace.items.clear;
  TraceErr.Clear;
 
  
    // si cheque dejeuner et code client / etabl ==> un ficbier final / etablissement
 if ((fournisseur = '004') and (Vh_Paie.PGFACTETABL = true)) then
 begin
      Q := Opensql ('SELECT ETB_TICKLIVR FROM ETABCOMPL', true);
      While not Q.EOF do
      begin
        codeclient := Q.Findfield('ETB_TICKLIVR').asstring;
        if codeclient <> '' then
        Conffichierfinal(codeclient);
      Q.Next;
      end;
  Ferme(Q);
 end
 else
 begin
 codeclient := '';
 Conffichierfinal(codeclient);
 end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 11/06/2003
Modifié le ... : 11/06/2003
Description .. : procédure ConfFichierFinal
Suite ........ : Lance la fiche de saisie de la commande de stock si la
Suite ........ : case à cocher = True
Suite ........ : Lance la confection du fichier final
Suite ........ : Copie le fichier généré sous son nom et sa destination
Suite ........ : définitifs (disquette, repertoire....)
Mots clefs ... : PAIE; CDETICKETRESTAU
*****************************************************************}
// pt4 procedure TOF_CDETICKETRESTAU.ConfFichierFinal(Sender: TObject);
procedure TOF_CDETICKETRESTAU.ConfFichierFinal(codeclient : string);   // pt4

var
  Pan                                   : TPageControl;
  Tbsht                                 : TTabSheet;
  CodeRetour                            : boolean;
  FFichierFinal                         : TextFile;
//PT1-4  Dir,FichierFinal,Supp,HuitZero,St     : string;
  Dir,FichierFinal,HuitZero,St          : string;
  ListeJ                                : HTStrings;
  rep                                   : integer;
  Q                                     : TQuery; // PT1-1


begin
  CodeRetour := True;
  Supp := GetControlText('SUPPORTEMIS');
  HuitZero := '00000000';

  St := DateDeb.Text+';'+DateFin.Text;
  if (GetControlText('STOCK') = 'X') then
    AglLanceFiche ('PAY','TICKETSTOCK', '', '' , St );

  Pan := TPageControl (GetControl ('PANELPREP'));
  Tbsht := TTabSheet (GetControl ('TBSHTTRACE'));
//  Trace := TListBox (GetControl ('LSTBXTRACE'));
//  TraceErr := TListBox (GetControl ('LSTBXERROR'));
//  Trace.Clear;

//  TraceErr.Clear;
  if (Trace = NIL) or (TraceErr = NIL) then
  begin
    PGIBox ('La confection du fichier de commande ne peut pas être lancée', 'Les composants trace ne sont pas disponibles');
    exit;
  end;
// d PT1-1
  St := 'SELECT * '+
        'FROM CDETICKETS '+
        'WHERE PRT_DATEDEBUT >= "'+USDateTime(StrToDate(Datedeb.text))+'" '+
        'AND PRT_DATEFIN <= "'+USDateTime(StrToDate(Datefin.text))+'"';
  Q := OpenSql (St, TRUE);
  if not Q.EOF then
  begin
    CodeRetour := True;
  end
  else
  begin
    PGIBox ('Aucune ligne de commande n''a été saisie pour cette période ', 'Confection impossible');
    TraceErr.Items.Add ('Aucune ligne de commande n''a été saisie pour cette période');
    CodeRetour := False;
  end;
  Ferme(Q);
// f PT1-1

  Dir :=  Trim(GetParamSocSecur('SO_PGREPERTTICKET',''));

  if not DirectoryExists(Dir) then
  begin
    PGIBox ('Le répertoire '+Dir+' est inexistant', 'Confection impossible');
    TraceErr.Items.Add ('Le répertoire '+Dir+' est inexistant');
    CodeRetour := False;
  end;

if CodeRetour = True then
begin
    Trace.Items.Add ('Commande de titres du '+DateCde.text);
    Trace.Items.Add ('Pour la prériode du '+DateDeb.text+
                     ' au '+DateFin.text);
    Trace.Items.Add ('');


  // pt4  OuvreFichierFinal(FFichierFinal, FichierFinal, Dir, CodeRetour);
   OuvreFichierFinal(FFichierFinal, FichierFinal, Dir, CodeRetour, codeclient); // pt4
    if (CodeRetour = False) then
    begin
      TraceErr.Items.Add('Fichier inaccessible : '+FichierFinal);
    end
    else
    begin
     CreatLigneCdeFinale(Dir,FFichierFinal, FichierFinal, CodeRetour, codeclient);
     CloseFile (FFichierFinal);
    end;
    if (CodeRetour = False) then
    begin
      TraceErr.Items.Add('');
      TraceErr.Items.Add('Confection du fichier abandonnée');
      TraceErr.Items.Add('Anomalie en mise à jour : '+FichierFinal);
      if (FileExists(FichierFinal)) then
        DeleteFile(PChar(FichierFinal));
    end;

  end;

  if TraceErr.items.Count < 1 then
  begin
    if (Supp = 'DTK') then
      {disquette}
      Dir := 'A:';

    if (Supp = 'TRA') then
      {répertoire de travail}
      Dir := VH_Paie.PGCheminEagl;

    if (Supp <> 'TEL') and (Supp <> 'DIS') then
    begin
     if not DirectoryExists(Dir) then
     begin
       if (Supp <> 'DTK') then
       begin
         PGIBox ('Le répertoire '+Dir+' est inexistant', 'Création fichier impossible');
         TraceErr.Items.Add ('Le répertoire '+Dir+' est inexistant');
         CodeRetour := False;
       end
       else
         while not DirectoryExists(Dir) do
         begin
           rep  := PGIAskCancel ('Le lecteur '+Dir+' est inaccessible. '+
                                 'Insérez une disquette');
           if (rep = mrCancel) then
           begin
             TraceErr.Items.Add ('Le lecteur '+Dir+' est inaccessible.');
             CodeRetour := False;
             break;
           end;
         end;
     end;
    end;
    if (CodeRetour = True) then
    begin
//d PT1-2
      if (Fournisseur = '001') then
      begin
      // SODEXHO
        CopyFile(PChar(FichierFinal),
                 PChar(Dir+
                '\'+
                 Copy(HuitZero,1,(8-length(Trim(VH_Paie.PgCodeRattach))))+
                 Trim(VH_Paie.PgCodeRattach)+
                 '.st1'),
                 False);
      end
      else
      if (Fournisseur = '002') then // PT2
      begin
      // NATEXIS
        CopyFile(PChar(FichierFinal),
                 PChar(Dir+
                '\'+
                'E'+
                 Copy(HuitZero,1,(7-length(Trim(VH_Paie.PgCodeRattach))))+
                 Trim(VH_Paie.PgCodeRattach)+
                 '.TXT'),
                 False);
// d PT2
      end
      else
      if (Fournisseur = '003') then
      begin
     // ACCOR
        CopyFile(PChar(FichierFinal),
                 PChar(Dir+
                '\'+
                 Copy(HuitZero,1,(6-length(Trim(VH_Paie.PgCodeRattach))))+
                 Trim(VH_Paie.PgCodeRattach)+
                 '.CSV'),
                 False);
// f PT2
    // pt3  end;
    // deb pt3
       end
       else
       if (fournisseur = '004') then
       begin
       // Chèque déjeuner
       if codeclient = '' then
         CopyFile(PChar(FichierFinal),
                 PChar(Dir+
                '\'+
                 Copy(HuitZero,1,(5-length(Trim(VH_Paie.PgCodeRattach))))+
                 Trim(VH_Paie.PgCodeRattach)+
                 '.prn'),
                 False)
         else
          CopyFile(PChar(FichierFinal),
                 PChar(Dir+
                '\'+
                 Copy(HuitZero,1,(5-length(Trim(codeclient))))+
                 Trim(codeclient)+
                 '.prn'),
                 False);

       end;
       // fin pt3
   //f PT1-2

      Trace.Items.Add ('');

// d PT1-2
      if (Fournisseur = '001') then
      begin
      // SODEXHO
        Trace.Items.Add ('Fichier créé : '+
                         Dir+
                         '\'+
                         Copy(HuitZero,1,(8-length(Trim(VH_Paie.PgCodeRattach))))+
                         Trim(VH_Paie.PgCodeRattach)+
                         '.st1');
      end
      else
      if (Fournisseur = '002') then //PT2
      begin
      // NATEXIS
        Trace.Items.Add ('Fichier créé : '+
                         Dir+
                         '\'+
                         'E'+
                         Copy(HuitZero,1,(7-length(Trim(VH_Paie.PgCodeRattach))))+
                         Trim(VH_Paie.PgCodeRattach)+
                         '.TXT');
// d  PT2
      end
      else
      if (Fournisseur = '003') then
      begin
      // ACCOR
        Trace.Items.Add ('Fichier créé : '+
                         Dir+
                         '\'+
                         Copy(HuitZero,1,(6-length(Trim(VH_Paie.PgCodeRattach))))+
                         Trim(VH_Paie.PgCodeRattach)+
                         '.CSV');
// f PT2
    // pt3  end;
     // deb pt3
      end
      else
      if (Fournisseur = '004') then
      begin
      // chèque déjeuner
        if codeclient = ''
        then
         Trace.Items.Add ('Fichier créé : '+
                         Dir+
                         '\'+
                         Copy(HuitZero,1,(5-length(Trim(VH_Paie.PgCodeRattach))))+
                         Trim(VH_Paie.PgCodeRattach)+
                         '.prn')
        else
         Trace.Items.Add ('Fichier créé : '+
                         Dir+
                         '\'+
                         Copy(HuitZero,1,(5-length(Trim(codeclient))))+
                         Trim(codeclient)+
                         '.prn');
      end;
      // fin pt3
// f PT1-2

      Trace.Items.Add ('Fin de traitement');
      Trace.Items.Add ('-----------------------------------------------------');
    end;
  end;

  if (Supp = 'TEL') and (CodeRetour = True) then
  begin
     ListeJ:=HTStringList.Create ;
     ListeJ.Add ('Veuillez trouver ci-joint notre commande de titres restaurant du '+DateCde.text);
     // pt3
     if (fournisseur = '004') then
     begin
      if codeclient = ''
      then
        ListeJ.Add ('correspondant au code client '+VH_Paie.PgCodeRattach)
      else
        ListeJ.Add ('correspondant au code client '+codeclient);

       ListeJ.Add ('pour une quantité totale de '+ inttostr(Wtotalticket) +' chèques');
     end;

     ListeJ.Add ('Cordialement');
// d PT1-2
     if (Fournisseur = '001') then
     begin
     // SODEXHO
     SendMail ('Commande de titres restaurant',
               '',
               '',
               ListeJ,
               Dir+'\'+
               Copy(HuitZero,1,(8-length(Trim(VH_Paie.PgCodeRattach))))+
               Trim(VH_Paie.PgCodeRattach)+
               '.st1',
               FALSE);
     end
     else
     if (Fournisseur = '002') then // PT2
     begin
     // NATEXIS
     SendMail ('Commande de titres restaurant',
               '',
               '',
               ListeJ,
               Dir+'\'+
               'E'+
               Copy(HuitZero,1,(7-length(Trim(VH_Paie.PgCodeRattach))))+
               Trim(VH_Paie.PgCodeRattach)+
               '.TXT',
               FALSE);
// d PT2
     end
     else
     if (Fournisseur = '003') then
     begin
     // ACCOR
     SendMail ('Commande de titres restaurant',
               '',
               '',
               ListeJ,
               Dir+'\'+
               Copy(HuitZero,1,(6-length(Trim(VH_Paie.PgCodeRattach))))+
               Trim(VH_Paie.PgCodeRattach)+
               '.CSV',
               FALSE);
// f PT2
  //   end;
  // deb pt3
    end
    else
    if (fournisseur = '004') then
    begin
    // Chèque déjeuner
      if codeclient = ''
      then
      SendMail ('Commande de titres restaurant',
               '',
               '',
               ListeJ,
               Dir+'\'+
               Copy(HuitZero,1,(5-length(Trim(VH_Paie.PgCodeRattach))))+
               Trim(VH_Paie.PgCodeRattach)+
               '.prn',
               FALSE)
      else
      SendMail ('Commande de titres restaurant',
               '',
               '',
               ListeJ,
               Dir+'\'+
               Copy(HuitZero,1,(5-length(Trim(codeclient))))+
               Trim(codeclient)+
               '.prn',
               FALSE);
      end;
  // fin pt3
// f PT1-2

     ListeJ.Clear ; ListeJ.Free ;
  end;

  if (CodeRetour = False) then
  begin
    TraceErr.Items.Add('');
    TraceErr.Items.Add('Création du fichier abandonnée');
    if (FileExists(FichierFinal)) then
      DeleteFile(PChar(FichierFinal));
  end;
  if TraceErr.items.Count >= 1 then
  begin
    Trace.Items.Add ('Fin du traitement, consultez vos anomalies');
    Trace.Items.Add ('-------------------------------------------------------');
  end;

  Pan.ActivePage := Tbsht;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 11/06/2003
Modifié le ... : 11/06/2003
Description .. : Procédure OuvreFichierFinal
Suite ........ : Ouverture, sur le répertoire défini au niveau des paramètres 
Suite ........ : société destinés aux tickets restaurant, du fichier final de 
Suite ........ : commande nommé Code client de rattachement (complété avec  0 à
Suite ........ : gauche)+date de la commande + .st1
Suite ........ : Ce fichier sera conservé
Mots clefs ... : PAIE; CDETICKETRESTAU
*****************************************************************}
// pt4 procedure TOF_CDETICKETRESTAU.OuvreFichierFinal(var FFichierFinal : TextFile; var FichierFinal : string; Dir : string; var CodeRetour : boolean);
procedure TOF_CDETICKETRESTAU.OuvreFichierFinal(var FFichierFinal : TextFile; var FichierFinal : string; Dir : string; var CodeRetour : boolean; codeclient : string);   // pt4

begin
  CodeRetour := True;

  if not DirectoryExists(Dir) then
    if not ForceDirectories(Dir) then
    begin
      PGIInfo('Impossible de créer le répertoire de travail '+Dir,
              'Titres restaurant');
      TraceErr.Items.Add ('Impossible de créer le répertoire de travail '+Dir);
      CodeRetour := False;
    end;

  if (CodeRetour = True) then
  begin
    if codeclient = '' then // pt4
    FichierFinal := Dir+'\'+VH_Paie.PgCodeRattach+FormatDateTime('ddmmyyyy',StrToDate(DateCde.text))+'.st1'
    else   // pt4
    FichierFinal := Dir+'\'+codeclient+FormatDateTime('ddmmyyyy',StrToDate(DateCde.text))+'.st1'; // pt4

    if (FileExists(FichierFinal)) then
      DeleteFile(PChar(FichierFinal));

    AssignFile(FFichierFinal, FichierFinal);
    {$i-} ReWrite(FFichierFinal); {$i+}
    if IoResult <> 0 then
      CodeRetour := False;
  end;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 11/06/2003
Modifié le ... : 11/06/2003
Description .. : procédure CreatLigneCdeFinale
Suite ........ : Parcours du répertoire de travail. Pour chaque fichier 
Suite ........ : Pxxxxxxxx.st1, si la date de commande correspond à celle 
Suite ........ : de la fiche. Si ce n'est pas le cas le fichier n'est pas traité.
Suite ........ : Mise à jour de la table RECAPCDETICKET.
Mots clefs ... : PAIE; CDETICKETRESTAU
*****************************************************************}
// pt4 procedure TOF_CDETICKETRESTAU.CreatLigneCdeFinale(Dir :string;var FFichierFinal : TextFile; FichierFinal : string;var CodeRetour : boolean);
procedure TOF_CDETICKETRESTAU.CreatLigneCdeFinale(Dir :string;var FFichierFinal : TextFile; FichierFinal : string;var CodeRetour : boolean; codeclient : string); // pt4

var
  sr                                                            : TsearchRec;
  FileAttrs, ret, NoErreur, NbT, TotTicket, NbCarnet            : integer;
  TTotTicket, TNbCarnet                                         : integer;
  TotValFaciale, TTotValFaciale                                 : double;
  RaisonSociale, HuitZero, blanc                                : string;
  FLect                                                         : TextFile;
  SIn, SOut                                                     : string;
  TOB_Recap, TOB_RecapFille                                     : TOB;
  Periode                                                       : string; //PT1-3
  NbEnreg                                                       : integer; // PT1-4
  st , VFaciale                                                 : string; //PT2
  DateLivraison                                                 : string; //PT2

begin
  blanc := StringOfChar(' ',61);
  HuitZero := '00000000';
  CodeRetour := True;
  TotValFaciale := 0;
  TTotValFaciale := 0;
  TTotTicket := 0;
  TNbCarnet := 0;
  TotTicket := 0;
  NbCarnet := 0;
  Wtotalticket := 0; // pt3
//d PT1-3
  Periode := Copy(Datedeb.text,9,2)+Copy(Datedeb.text,4,2)+Copy(Datedeb.text,1,2)+
             Copy(Datefin.text,9,2)+Copy(Datefin.text,4,2)+Copy(Datefin.text,1,2) ;
//f PT1-3

  FileAttrs := 0;
  FileAttrs := FileAttrs + faAnyFile;
//PT1-3  ret := FindFirst(Dir+'\P*.st1',FileAttrs,sr);

 if codeclient = '' then     // pt4
  ret := FindFirst(Dir+'\P*'+Periode+'.st1',FileAttrs,sr)
  else  ret := FindFirst(Dir+'\P'+ codeclient+Periode+'.st1',FileAttrs,sr); // pt4

  if (ret = 0) then
  begin
// d PT1-2
    if (Fournisseur = '001') then
    begin
    // SODEXHO
// f PT1-2
      {Enregistrement d'en-tête}
      RaisonSociale := GetParamSocSecur('SO_LIBELLE','');
      Writeln(FFichierFinal,copy(HuitZero,1,(8-length(Trim(VH_Paie.PgCodeRattach))))+
                        Trim(Vh_Paie.PgCodeRattach)+
                        FormatDateTime('ddmmyyyy',StrToDate(DateCde.text))+
                        PgUpperCase(RaisonSociale)+
                        Copy(blanc,1,61-length(RaisonSociale))+
                        '  '+
                        '1');
    end; //PT1-2

    NbEnreg := 0;       //PT1-4
    while (ret = 0) do
    begin
      { Traitement d'un pré-fichier }
      AssignFile(Flect,Dir+'\'+ sr.Name);
      {$i-} Reset(FLect); {$i+}
      NoErreur := IoResult;
      if NoErreur<>0 then
      begin
        PGIBox('('+IntToStr(NoErreur)+') Fichier inaccessible : '+Dir+'\'+sr.Name,
               'Fichier non traité');
        closeFile(FLect);
        CodeRetour := False;
        exit ;
      end;

      while not eof(FLect) do
      begin
        {$i-} Readln (FLect,Sin); {$i+}
        NoErreur := IoResult;
        if NoErreur<>0 then
        begin
          PGIBox('('+IntToStr(NoErreur)+') Erreur de lecture du fichier : '+Dir+'\'+Sr.Name,'Déclaration non traitée'); // PT4-3
          closeFile(FLect);
          CodeRetour := False;
          Exit ;
        end;
// d PT1-2
        if (Fournisseur = '001') then
        begin
        // SODEXHO
// f PT1-2
          if (copy(Sin,80,1) = '2') then
          { seul les enregistrements de fabrication sont traités}
          begin
            SOut := Sin;
            if (SOut <> '') then
            begin
              {$i-} Writeln (FFichierFinal,SOut); {$i+}
              NbT := StrToInt(Copy(Sin,19,2));
              TotValFaciale := TotValFaciale +((StrToInt(Copy(Sin,21,5)))/100) * NbT;
              TotTicket := TotTicket + NbT;
              if ((NbT div 25) * 25) < NbT  then
                NbCarnet := NbCarnet + ((NbT div 25) + 1)
              else
                NbCarnet := NbCarnet + (NbT div 25);
            end;
          end;
// d PT1-2
        end
        else
        if (Fournisseur = '002') then  //PT2
        begin
        // NATEXIS
          if (copy(Sin,99,1) = '2') then
         { seul les enregistrements de fabrication sont traités}
          begin
            SOut := copy(Sin,1,98);
            if (SOut <> '') then
            begin
              {$i-} Writeln (FFichierFinal,SOut); {$i+}
              NbT := StrToInt(Copy(Sin,20,5));
              Sin := Copy(Sin,1,26)+DecimalSeparator+Copy(Sin,28,72);
              TotValFaciale := TotValFaciale +((StrTofloat(Copy(Sin,25,5)))) * NbT;
              TotTicket := TotTicket + NbT;
              NbCarnet := NbCarnet + StrToInt(Copy(Sin,15,3));
              NbEnreg := NbEnreg + 1; // PT1-4
            end;
          end;
// d PT2
        end
        else
        if (Fournisseur = '003') then 
        begin
        // ACCOR
          if (copy(Sin,1,3) = '50;') then
         { seul les enregistrements de fabrication sont traités}
          begin
            SOut := copy(Sin,1,length(Sin));
            if ((DateLivraison <> '') and (copy(sr.name,1,6) = 'PSTOCK')) then
              SOut := SOut+';'+DateLivraison;

            if (SOut <> '') then
            begin
              {$i-} Writeln (FFichierFinal,SOut); {$i+}
              st := ReadTokenSt(Sin);
              st := ReadTokenSt(Sin);
              st := ReadTokenSt(Sin);
              st := ReadTokenSt(Sin);
              st := ReadTokenSt(Sin);
              st := ReadTokenSt(Sin);
              st := ReadTokenSt(Sin);
              st := ReadTokenSt(Sin);
              st := ReadTokenSt(Sin);
              VFaciale := ReadTokenSt(Sin);
              st := ReadTokenSt(Sin);
              NbT := StrToInt(ReadTokenSt(Sin));
              TotValFaciale := TotValFaciale +((StrTofloat(VFaciale))/100) * NbT;
              TotTicket := TotTicket + NbT;
              NbCarnet := NbCarnet + 1;
              if (copy(sr.name,1,6) <> 'PSTOCK') then
              begin
                DateLivraison := '';
                if (GetParamSocSecur ('SO_PGTKDATELIVR', FALSE) = TRUE) then
                begin
                 st := ReadTokenSt(Sin);
                 st := ReadTokenSt(Sin);
                 st := ReadTokenSt(Sin);
                 st := ReadTokenSt(Sin);
                 st := ReadTokenSt(Sin);
                 DateLivraison := ReadTokenSt(Sin);
                end;
              end;
            end;
          end;
//  f PT2
    // pt3    end;
    // deb pt3
        end
        else
        if (Fournisseur = '004') then
        begin
        // Chèque déjeuner
          if (copy(Sin,1,1) = '4') then
         { seul les enregistrements de fabrication sont traités}
          begin
            if (Sin <> '') then
            begin
              {$i-} Writeln (FFichierFinal,Sin); {$i+}
              NbT := StrToInt(Copy(Sin,56,2));
              TotTicket := TotTicket + NbT;
              TotValFaciale := TotValFaciale +((StrTofloat(Copy(Sin,59,4)))/100) * NbT;
              Nbcarnet := Nbcarnet + 1;
            end;
          end;
        end;
        // fin pt3
// f PT1-2
        NoErreur := IoResult;
        if NoErreur<>0 then
        begin
          PGIBox('('+IntToStr(NoErreur)+') Erreur d''écriture du fichier : '+Dir+'\'+sr.Name,'Déclaration non traitée'); // PT4-3
          closeFile(FLect);
          CodeRetour := False;
          Exit ;
        end;
        if (Copy(Sin,80,1) = '1') then
        { Sur l'enregistrement d'en-tête, vérification date de Cde OK}
        begin
          if (Copy(Sin,11,6) <> Copy(FormatDateTime('ddmmyyyy',StrToDate(DateCde.text)),3,6)) then break;

          if not assigned(TOB_Recap) then
            TOB_Recap := TOB.create('le récapitulatif',NIL,-1);

          TOB_RecapFille := TOB.Create('RECAPCDETICKET', TOB_Recap,-1);
          TOB_RecapFille.PutValue('PHT_DATEDEBUT',StrToDate(Datedeb.text));
          TOB_RecapFille.PutValue('PHT_DATEFIN',StrToDate(Datefin.text));
          TOB_RecapFille.PutValue('PHT_DATECDE',StrToDate(DateCde.text));
          TOB_RecapFille.PutValue('PHT_AUTEUR',GetControlText('AUTEUR'));
          TOB_RecapFille.PutValue('PHT_CODECLIENT',Copy(Sin,1,8));
//          TOB_RecapFille.PutValue('PHT_FRAISGESTION',Arrondi(StrToFloat(GetParamSoc('SO_PGFRAISGESTION')),5));
          TOB_RecapFille.PutValue('PHT_NOMFICH',FichierFinal);
        end;
        if (Copy(Sin,80,1) = '3') then
        begin
          TTotValFaciale := TTotValFaciale + TotValFaciale;
          TTotTicket := TTotTicket + TotTicket;
          TNbCarnet := TNbCarnet + NbCarnet;
          TOB_RecapFille.PutValue('PHT_NBTICKET',IntToStr(TTotTicket));
          TOB_RecapFille.PutValue('PHT_NBCARNET',IntToStr(TNbCarnet));
          TOB_RecapFille.PutValue('PHT_CUMVFACIALE',Arrondi(TTotValFaciale,5));
          TOB_RecapFille.PutValue('PHT_FRAISGESTION',
                (Arrondi(Valeur(GetParamSocSecur('SO_PGFRAISGESTION',0.00000)),5)*
                 Arrondi(TTotValFaciale,5))/100);
          TotTicket := 0;
          NbCarnet := 0;
          TotValFaciale := 0;
          end;
      end;
// d PT1-2
 // PT3   if (Fournisseur = '002') then
  if ((Fournisseur = '002') or (fournisseur = '004')) then    // pt3
      begin
      // NATEXIS  ou chèque déjeuner
        TTotValFaciale := TTotValFaciale + TotValFaciale;
        TNbCarnet := TNbCarnet + NbCarnet;
        TTotTicket := TTotTicket + TotTicket;
      end;
// f PT1-2

      CloseFile (FLect);
      ret := FindNext(sr);
    end;  {fin while (ret = 0)}
    sysutils.FindClose(sr);

    if (TTotTicket = 0) then
    begin
      PGIBox ('Attention : Aucun titre commandé', 'Confection impossible');
      TraceErr.Items.Add ('!! Attention : Aucun titre commandé');
      CodeRetour := False;
    end;
    if (CodeRetour = True) then
    begin
// d PT1-2
      if (Fournisseur = '001') then
      begin
        {Enregitrement de fin de fichier}
        writeln(FFichierFinal,'********'+
                              ColleZeroDevant(TTotTicket,6)+
                              ColleZeroDevant(TNbCarnet,6)+
                              copy(blanc,1,59)+
                              '3');
      end;
//f PT1-2

      Trace.Items.Add ('Nombre de carnets commandés : '+
                       IntToStr(TNbCarnet));
      Trace.Items.Add ('Nombre de titres commandés  : '+
                       IntToStr(TTotTicket));
      Trace.Items.Add ('Montant valeurs faciales    : '+
                       FormatFloat('#0.00',TTotValFaciale));

      Wtotalticket := Ttotticket; // pt3

      try
        BeginTrans ;
        if assigned (TOB_Recap) then
        begin
          TOB_Recap.SetAllModifie(True);
          TOB_Recap.InsertOrUpdateDB(False);
        end;
        CommitTrans;
      except
        RollBack;
        PGIBox('! Erreur maj table RECAPCDETICKET','');
      end ;
    end;   { fin if (CodeRetour = True)}
  end; { fin if (ret = 0)}
  { suppression du fichier de commande de stock s'il existe}
// PT1-3 if (FileExists(Dir+'\PSTOCK.st1')) then
// PT1-3 DeleteFile(PChar(Dir+'\PSTOCK.st1'));
  if (FileExists(Dir+'\PSTOCK'+Periode+'.st1')) then
    DeleteFile(PChar(Dir+'\PSTOCK'+Periode+'.st1'));

  if (Fournisseur = '002') then
  begin
  // NATEXIS
      Tob_RecapFille.AddChampSupValeur('NBENREG',NbEnreg);
      Tob_RecapFille.AddChampSupValeur('NOCLIENT',Vh_Paie.PgCodeRattach);
      if (Supp = 'TEL') then
        Tob_RecapFille.AddChampSupValeur('MAIL','X')
      else
      Tob_RecapFille.AddChampSupValeur('MAIL','-');

    Tob_RecapFille.AddChampSupValeur('COURRIER','-');
    Tob_RecapFille.AddChampSupValeur('NOMFIC','E'+
                 Copy(HuitZero,1,(7-length(Trim(VH_Paie.PgCodeRattach))))+
                 Trim(VH_Paie.PgCodeRattach)+
                 '.TXT');
  end;
  if (Fournisseur = '002') then
    LanceEtatTOB('E','PAY','PNA',Tob_Recap,True,False,False,Pages,'','',False);  //   PT1-4
  FreeAndNil (TOB_Recap); // PT1-4

end;

Initialization
  registerclasses ( [ TOF_CDETICKETRESTAU ] ) ;
end.
