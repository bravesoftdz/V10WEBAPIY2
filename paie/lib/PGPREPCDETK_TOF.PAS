{***********UNITE*************************************************
Auteur  ...... : PAIE MF
Créé le ...... : 16/05/2003
Modifié le ... : 06/06/2003
Description .. : Source TOF de la FICHE : PGPREPCDETK ()
Suite ........ : Lancement du pré-fichier de commande de tickets
Suite ........ : restaurant
Mots clefs ... : TOF;PGPREPCDETK
*****************************************************************}
{
   PT1    MF    14/02/2005  V_6.0  1- Le nom du pré-fichier de commande contient
                                   les dates de prériode (--> autant de pré-fichiers
                                   que de périodes)
                                   2- Mise en place de la confection du fichier
                                   pour NATEXIS
   PT2    MF    14/02/2005  V_6.0  Correction traitement qd la dernière ligne de
                                   commande = 0 ticket  (bouclait)
   PT3    MF    25/02/2005  V_6.0  Ne sont sélectionnées pour la cde que les lignes
                                   ayant un code distribution existant
   PT4    MF    24/03/2005  V_60   Correction Qd nbre tickets = 0 les salariés
                                   suivants ne faisaient pas partie de la cde.
   PT5    MF    06/12/2006  V_700  FQ 13173 : correction fichier incomplet qd dernière ligne à 0
                                   + correction division par 0 qd toutes les lignes à 0
   PT6    MF    22/01/2007  V_720  Nouveau fournisseur : ACCOR
   PT7    NA    29/05/2008  V_850  Nouveau fournisseur : CHEQUE DEJEUNER

}
Unit PGPREPCDETK_TOF ;

Interface

Uses
     {$IFDEF VER150}
     Variants,
     {$ENDIF}
     StdCtrls,
//unused     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
//unused     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     HTB97,
//unused     PgOutils,
     ParamSoc,
     UTOB,
     FileCtrl,
     EntPaie,
     PgOutils2;

Type
  TOF_PGPREPCDETK = Class (TOF)
    private
    DateDeb,DateFin,DateCde            : THEdit;
    BtnLance                           : TToolbarButton97;
    CodeClient, RaisonSociale          : string;
    Trace,TraceErr                     : TListBox;
    Fournisseur                        : string; // PT1-2

    procedure ConfPreFichier(Sender: TObject);
// PT1-1    procedure OuvrePreFichier(var FPrefichier : TextFile; var Prefichier : string;var CodeRetour : boolean);
    procedure OuvrePreFichier(var FPrefichier : TextFile; var Prefichier : string;var CodeRetour : boolean; const Periode : string);
    procedure CreatLigneCde(TobCde : TOB; var FPrefichier : TextFile; var Prefichier : string;var CodeRetour : boolean);
    procedure Majclient(Sender: TObject); // pt7

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

procedure TOF_PGPREPCDETK.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGPREPCDETK.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGPREPCDETK.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGPREPCDETK.OnLoad ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE MF
Créé le ...... : 06/06/2003
Modifié le ... : 06/06/2003
Description .. : procédure OnArgument
Suite ........ : présentation de la période pour laquelle le traitement est 
Suite ........ : réalisé. Il s'agit de la période en cours, cette période est 
Suite ........ : modifiable.
Suite ........ : 
Mots clefs ... : PAIE; PGPREPCDETK
*****************************************************************}
procedure TOF_PGPREPCDETK.OnArgument (S : String ) ;
var
  DebPer,FinPer,ExerPerEncours              : String;
  OkOk                                      : Boolean;
  periode                                   : string; // PT1-1
  CHQDEJclient                              : THEDIT; // pt7
begin
  if (VH_Paie.PgCodeClient = '') or (VH_Paie.PgCodeClient = NULL) then
  begin
    PgiBox('le code client n''est pas renseigné',
           'Pré-fichier de commande');
    exit;
  end;
//d PT1-2
  Fournisseur := GetParamSocSecur('SO_PGTYPECDETICKET','');
  if (length(VH_Paie.PgCodeClient) > 7) and
     (Fournisseur = '002') then
  begin
    PgiBox('le code client facturé doit être au maximum de 7 caractères'+#10#13+
           ' A modifier dans les paramètres société' ,
           'Pré-fichier de commande');
    exit;
  end;
// f PT1-2

  Inherited ;
  DateDeb := ThEdit(getcontrol('DATEDEB'));
  DateFin := ThEdit(getcontrol('DATEFIN'));
  DateCde := ThEdit(getcontrol('DATECDE'));
  OkOk := RendPeriodeEnCours (ExerPerEncours,DebPer,FinPer);
  if OkOk then
  begin
   if  DateDeb <> NIL   then DateDeb.text:= DebPer;
   if  DateFin <> NIL   then DateFin.text:= FinPer;
  end;
//d PT1-1
  Periode := Copy(Datedeb.text,9,2)+Copy(Datedeb.text,4,2)+Copy(Datedeb.text,1,2)+
             Copy(Datefin.text,9,2)+Copy(Datefin.text,4,2)+Copy(Datefin.text,1,2) ;
//f PT1-1

  if DateCde <> NIl then DateCde.text := DateToStr(Date);
  
  BtnLance:=TToolbarButton97 (GetControl ('BLANCE'));
  if BtnLance<>NIL then
    BtnLance.OnClick := ConfPreFichier;

  CodeClient := VH_Paie.PgCodeClient;
  RaisonSociale := GetParamSocSecur('SO_LIBELLE','');
// PT1-1         SetControlText ('NOMFIC','P'+CodeClient+'.st1');

  if ((fournisseur = '004') and ( VH_Paie.PGFACTETABL = true ))then   // pt7
  setcontroltext('NOMFIC', '')                                         // pt7
  else                                                                 // pt7
  SetControlText ('NOMFIC','P'+CodeClient+Periode+'.st1');
// d PT6
// Pour ACCOR possibilité de renseigner une date de livraison
   if (GetParamSocSecur ('SO_PGTKDATELIVR', FALSE) = TRUE) then
   begin
     SetControlVisible('DATELIVR', TRUE);
     SetControlEnabled('DATELIVR', TRUE);
     SetControlVisible('LDATELIVR', TRUE);
   end;
// f PT6
// deb pt7 si chèque déjeuner : saisi du code client si code client par établissement
  if ((fournisseur = '004') and ( VH_Paie.PGFACTETABL = true )) then
  begin
     SetControlVisible('CODECLIENT', TRUE);
     SetControlEnabled('CODECLIENT', TRUE);
     SetControlVisible('T_CODECLIENT', TRUE);
     CHQDEJclient := ThEdit(getcontrol('CODECLIENT'));
     if CHQDEJclient <> NIL then CHQDEJCLIENT.onexit := MAJclient;
  end;
// fin pt7
end ;

procedure TOF_PGPREPCDETK.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PGPREPCDETK.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGPREPCDETK.OnCancel () ;
begin
  Inherited ;
end ;

// deb pt7
{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 23/06/2008
Modifié le ... :   /  /    
Description .. : Si chèque déjeuner et facturation par établ :   préparation par code client
Suite ........ : formattage du nom du fichier
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGPREPCDETK.Majclient(Sender: TObject);
var
periode : string;
Begin
  Periode := Copy(Datedeb.text,9,2)+Copy(Datedeb.text,4,2)+Copy(Datedeb.text,1,2)+
             Copy(Datefin.text,9,2)+Copy(Datefin.text,4,2)+Copy(Datefin.text,1,2) ;
  codeclient := getcontroltext('CODECLIENT') ;
  if codeclient = '' then
  begin
   PgiBox('Le code client  doit être renseigné.', 'Pré-fichier de commande');
   setfocuscontrol('CODECLIENT');
   exit;
  end;

  if (length(codeclient) <> 5) then
  begin
   PgiBox('Le code client  doit être sur 5 caractères.', 'Pré-fichier de commande');
   setfocuscontrol('CODECLIENT');
   exit;
  end;
   SetControlText ('NOMFIC','P'+CodeClient+Periode+'.st1');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE MF
Créé le ...... : 11/06/2003
Modifié le ... : 11/06/2003
Description .. : procédure ConfPreFichier :
Suite ........ : confection du pré fichier de commande ( 1 pré fichier par
Suite ........ : code client )
Suite ........ : A partir de la table CDETICKETS préparation de la TOB
Suite ........ : TOBCde qui permet la mise à jour du pré fichier. Les
Suite ........ : lignes de commande sélectionnée sont celles
Suite ........ : correspondant aux dates de période. Mise à jour
Suite ........ : de la date de commande sur CDETICKETS
Mots clefs ... : PAIE; PGPREPCDETK
*****************************************************************}
procedure TOF_PGPREPCDETK.ConfPreFichier(Sender: TObject);
var
  TobCde                                : TOB;
  Q                                     : TQuery;
  Pan                                   : TPageControl;
  Tbsht                                 : TTabSheet;
  CdeExist, CodeRetour                  : boolean;
  FPreFichier                           : TextFile;
  PreFichier, St                        : string;
  Periode                               : string; // PT1-1
// d PT3
  StDistribution                               : string;
// f PT3
 
begin
  CdeExist := False;
  TobCde :=  NIL;

  Pan := TPageControl (GetControl ('PANELPREP'));
  Tbsht := TTabSheet (GetControl ('TBSHTTRACE'));
  Trace := TListBox (GetControl ('LSTBXTRACE'));
  TraceErr := TListBox (GetControl ('LSTBXERROR'));
  Trace.Clear;
  TraceErr.Clear;

// d PT6
// Contrôle Date de livraison renseignée
  if (GetParamSocSecur ('SO_PGTKDATELIVR', FALSE) = TRUE) then
  begin
    if (StrToDate(GetControlText('DATELIVR')) = IDate1900) then
    begin
      PgiBox('La date n''est pas renseignée','');
      SetFocusControl('DATELIVR');
      Exit ;
    end;
  end;
// f PT6

  if (Trace = NIL) or (TraceErr = NIL) then
  begin
    PGIBox ('La confection du Pré-fichier de commande ne peut pas être lancée', 'Les composants trace ne sont pas disponibles');
    exit;
  end;

// d PT3
  StDistribution := '';
  Q:= OpenSql('SELECT CC_ABREGE FROM CHOIXCOD '+
              'WHERE CC_TYPE="PCD"',True);
  While  not Q.eof do
  begin
    if (StDistribution = '') then
      StDistribution := ' AND (PRT_DISTRIBUTION = '
    else
      StDistribution := StDistribution + ' OR PRT_DISTRIBUTION = ' ;

    StDistribution := StDistribution+'"'+ Q.FindField('CC_ABREGE').AsString+'"';
    Q.Next;
  end;
  if (StDistribution <> '') then StDistribution := StDistribution + ')';
  ferme (Q);

// f PT3
  TobCde := TOB.Create ('La commande',NIL,-1);
  St := 'SELECT * '+
        'FROM CDETICKETS '+
        'WHERE PRT_DATEDEBUT >= "'+USDateTime(StrToDate(Datedeb.text))+'" '+
        'AND PRT_DATEFIN <= "'+USDateTime(StrToDate(Datefin.text))+'"';

// pt7
  if ((fournisseur = '004') and ( VH_Paie.PGFACTETABL = true )) then
  begin

    St := 'SELECT * '+
        'FROM CDETICKETS '+
        'WHERE PRT_DATEDEBUT >= "'+USDateTime(StrToDate(Datedeb.text))+'" '+
        'AND PRT_DATEFIN <= "'+USDateTime(StrToDate(Datefin.text))+'" '+
        'AND PRT_CODECLIENT = "'+codeclient+'"';
  end;
// fin pt7

// d PT3
  if (StDistribution <> '') then
    St := St + StDistribution ;
// f PT3
  Q := OpenSql (St, TRUE);
  if not Q.EOF then
  begin
    TobCde.LoadDetailDB ('COMMANDE TICKETS', '','',Q,FALSE,FALSE);
    CdeExist := True;
    TobCde.Detail.Sort ('PRT_SALARIE;PRT_DATEDEBUT;PRT_DATEFIN')
  end
  else
  begin
    PGIBox ('Aucune ligne de commande n''a été saisie pour cette période ', 'Confection impossible');
    TraceErr.Items.Add ('Aucune ligne de commande n''a été saisie pour cette période');
  end;
  Ferme(Q);
//d PT1-1
  Periode := Copy(Datedeb.text,9,2)+Copy(Datedeb.text,4,2)+Copy(Datedeb.text,1,2)+
             Copy(Datefin.text,9,2)+Copy(Datefin.text,4,2)+Copy(Datefin.text,1,2) ;
//f PT1-1

  if CdeExist = True then
  begin
    Trace.Items.Add ('Commande de titres du '+DateCde.text);
    Trace.Items.Add ('Pour la période du '+DateDeb.text+
                     ' au '+DateFin.text);
    Trace.Items.Add ('');
// PT1-1   OuvrePreFichier(FPrefichier, Prefichier, CodeRetour);
    OuvrePreFichier(FPrefichier, Prefichier, CodeRetour,Periode);
    if (CodeRetour = False) then
    begin
      TraceErr.Items.Add('Fichier inaccessible : '+Prefichier);
    end
    else
    begin
     CreatLigneCde(TobCde, FPreFichier, PreFichier, CodeRetour);
     CloseFile (FPreFichier);
    end;
    if (CodeRetour = False) then
    begin
      TraceErr.Items.Add('');
      TraceErr.Items.Add('Confection du pré-fichier abandonnée');
      TraceErr.Items.Add('Anomalie en mise à jour : '+Prefichier);
      if (FileExists(PreFichier)) then
        DeleteFile(PChar(Prefichier));
    end;
  end;

  if TraceErr.items.Count >= 1 then
    Trace.Items.Add ('Fin du traitement, consultez vos anomalies')
  else
  begin
    Trace.Items.Add ('');
    Trace.Items.Add ('Pré-fichier créé : '+PreFichier);
    Trace.Items.Add ('Fin de traitement');
    try
      BeginTrans ;
      St := 'UPDATE CDETICKETS SET PRT_DATECDE = "'+
            USDateTime(StrToDate(DateCde.text))+'", '+
            'PRT_DATEINTEG = "'+USDateTime(IDate1900)+'" '+
            'WHERE PRT_DATEDEBUT >= "'+USDateTime(StrToDate(Datedeb.text))+'" '+
            'AND PRT_DATEFIN <= "'+USDateTime(StrToDate(Datefin.text))+'"'+
            'AND PRT_NBTICKET > 0';
// d PT3
      if (StDistribution <> '') then
        St := St + StDistribution ;
// f PT3

      ExecuteSQL(St);

      CommitTrans;
    except
      RollBack;
      PGIBox('! Erreur maj table CDETICKETS','');
    end ;
  end;
  Pan.ActivePage := Tbsht;

  FreeAndNil(TobCde);

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE MF
Créé le ...... : 11/06/2003
Modifié le ... :   /  /
Description .. : Pocédure OuvrePreFichier
Suite ........ : Ouverture du pré fichier sur la répertoire paramétré au
Suite ........ : niveau des paramètres société. Annule et remplace le
Suite ........ : pré fichier déjaà existant pour ce code client.
Suite ........ : Nom du pré fichier : P + Code client + .st1
Mots clefs ... : PAIE; PGPREPCDETK
*****************************************************************}
// PT1-1 procedure TOF_PGPREPCDETK.OuvrePreFichier(var FPrefichier : TextFile; var Prefichier : string;var CodeRetour : boolean);
procedure TOF_PGPREPCDETK.OuvrePreFichier(var FPrefichier : TextFile; var Prefichier : string;var CodeRetour : boolean; const Periode :string);
var
  Dir                   : string;
begin
  CodeRetour := True;

  Dir :=  Trim(GetParamSocSecur('SO_PGREPERTTICKET',''));  
  if not DirectoryExists(Dir) then
    if (Dir = '') or (not ForceDirectories(Dir)) then
    begin
      PGIInfo('Impossible de créer le répertoire de travail '+Dir,
              'Titres restaurant');
      TraceErr.Items.Add ('Impossible de créer le répertoire de travail '+Dir);
      CodeRetour := False;
    end;

  if (CodeRetour = True) then
  begin
// on supprime le fichier ancienne version (sans dates de période)
    PreFichier := Dir+'\P'+CodeClient+'.st1';
    if (FileExists(PreFichier)) then
      DeleteFile(PChar(Prefichier));

// PT1-1    PreFichier := Dir+'\P'+CodeClient+'.st1';
    PreFichier := Dir+'\P'+CodeClient+Periode+'.st1';
    if (FileExists(PreFichier)) then
      DeleteFile(PChar(Prefichier));

    AssignFile(FPreFichier, PreFichier);
    {$i-} ReWrite(FPreFichier); {$i+}
    if IoResult <> 0 then
      CodeRetour := False;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE MF
Créé le ...... : 11/06/2003
Modifié le ... : 11/06/2003
Description .. : Procédure CreatLigneCde
Suite ........ : Mise à jour du pré fichier à partir de la TOBCde
Suite ........ : 1 enregitrement en-tête (1)
Suite ........ : x enregistrements de fabrication (2)
Suite ........ : 1 enregistrement de fin de fichieir (3)
Suite ........ : La quantité de tickets par ligne de fabrication est limitée à
Suite ........ : 99. Si pour un salarié la commande est > 99 plusieurs lignes
Suite ........ : seront générées.
Mots clefs ... : PAIE; PGPREPCDETK
*****************************************************************}
procedure TOF_PGPREPCDETK.CreatLigneCde(TobCde : TOB; var FPrefichier : TextFile; var Prefichier : string;var CodeRetour : boolean);
var
  TobCdeFille                                   : TOB;
  NoOrdre, NbT, NbLigne, NbTicket               : integer;
  TotTicket, NbCarnet                           : integer;
  blanc,HuitZero,VFaciale, PPatronale,SalPrec   : string;
  Huitblanc                                     : string; // pt7
  Distribution, Salarie, NomPrenom, Personnal   : string;
  WCodeClient, WRaisonSociale                   : string;
  SepDec                                        : Char; // PT1-2
  NbCarnetLigne,NbTicketCarnet,NbTotalLigne     : integer; //PT1-2
  NbCarnetW,NbCarnetTotal, Reste                : integer; //PT1-2
// d PT6
  QQ                                            : Tquery;
  NiveauInterne1, NiveauInterne2,EdNPCarnet     : String;
  EdNPTicket,EdRSCarnet,EdCPVCarnet             : string;
  CodePostal,Ville,PointdeLivraison,ANbt        : string;
  DateLivraison                                 : string;
  // f PT6
  i , rang , nbticketprec                       : integer;     // pt7
  abregedistrib                                 : string;      // pt7

begin
  TotTicket := 0;
  NbCarnet := 0;
  blanc := StringOfChar(' ',61);
  Huitblanc := StringOfChar(' ',8);   // pt7
  CodeRetour := True;
  HuitZero := '00000000';
  SalPrec := '';
  NbT := 0;
//d PT1-2
  if (Fournisseur = '001') then
  begin
  // SODEXHO
  if (length(Trim(CodeClient)) = 8) then
    WCodeClient := Trim(CodeClient)
  else
    WCodeClient := copy(HuitZero,1,(8-length(Trim(CodeClient))))+
                   Trim(CodeClient);
  end
  else
  if (Fournisseur = '002') then // PT6
  begin
  // NATEXIS
    if (length(Trim(CodeClient)) = 7) then
    WCodeClient := Trim(CodeClient)
  else
    WCodeClient := copy(HuitZero,1,(7-length(Trim(CodeClient))))+
                   Trim(CodeClient);
  end
// d PT6
  else
  if (Fournisseur = '003') then
  begin
  // ACCOR
    if (length(Trim(CodeClient)) = 6) then
    WCodeClient := Trim(CodeClient)
  else
    WCodeClient := copy(HuitZero,1,(6-length(Trim(CodeClient))))+
                   Trim(CodeClient);
 //pt7 end;
 // deb pt7
  end
  else
  if (Fournisseur = '004') then
  begin
  // CHEQUE DEJEUNER
   
     if (length(Trim(CodeClient)) = 5) then
    WCodeClient := Trim(CodeClient)
  else
    WCodeClient := copy(huitZero,1,(5-length(Trim(CodeClient))))+
                   Trim(CodeClient);
  end;
  // fin pt7
// f PT6
// f PT1-2

  if (length(RaisonSociale) = 61) then
    WRaisonSociale := PgUpperCase(RaisonSociale)
  else
    WRaisonSociale :=  PgUpperCase(RaisonSociale)+
                       Copy(blanc,1,61-length(RaisonSociale));

  {Enregistrement d'en-tête}
// d PT1-2
  if (Fournisseur ='001') then
  begin
  // SODEXHO
    Writeln(FPreFichier,WCodeClient+
                        FormatDateTime('ddmmyyyy',StrToDate(DateCde.text))+
                        WRaisonSociale+
                        '  '+
                        '1');
  end
  else
  if (Fournisseur ='002') then   // PT6
  begin
  // NATEXIS
    Writeln(FPreFichier,WCodeClient+
                        ' '+
                        FormatDateTime('ddmmyyyy',StrToDate(DateCde.text))+
                        WRaisonSociale+
                        '  '+
                        '1');
// d PT6
   end
   else
  if (Fournisseur ='003') then 
  begin
  // ACCOR
    Writeln(FPreFichier,WCodeClient+
                        '  '+
                        FormatDateTime('ddmmyyyy',StrToDate(DateCde.text))+
                        WRaisonSociale+
                        '  '+
                        '1');

// f PT6
//  end; pt3
// deb pt3
   end
   else
  if (fournisseur = '004') then
  begin
  // Chèque déjeuner
    Writeln(FPreFichier,WCodeClient+
                        '   '+
                        FormatDateTime('ddmmyyyy',StrToDate(DateCde.text))+
                        WRaisonSociale+
                        '  '+
                        '1');
  end;
// fin pt7
// f PT1-2

  {Enregistrement de fabrication}
  rang := 0; // pt7

  TobCdeFille := TobCde.FindFirst ([''],[''], True);
  if (TobCdeFille <> Nil) then
    SalPrec := TobCdeFille.GetValue('PRT_SALARIE');
  while (TobCdeFille  <> Nil) or ((TobCdeFille = Nil) and (SalPrec <> '')) do
  begin
    NoOrdre := 0;
    if (TobCdeFille  <> Nil) and
       (TobCdeFille.GetValue('PRT_SALARIE') = SalPrec) and
       (TobCdeFille.GetValue('PRT_NBTICKET') <> 0) then     // PT5
    begin
      NbT := NbT + TobCdeFille.GetValue('PRT_NBTICKET');

//d PT1-2
      if (Fournisseur = '001') or (Fournisseur = '003')then   // PT6
      begin
      // SODEXHO  ou ACCOR (en cts sur 5 car)
        if (length(FormatFloat('0',TobCdeFille.GetValue('PRT_VFACIALE')*100)) = 5) then
          VFaciale := FormatFloat('0',TobCdeFille.GetValue('PRT_VFACIALE')*100)
        else
          VFaciale := copy(HuitZero,1,(5-length(FormatFloat('0',TobCdeFille.GetValue('PRT_VFACIALE')*100))))+
                    FormatFloat('0',TobCdeFille.GetValue('PRT_VFACIALE')*100);
      end
      else
      if (Fournisseur = '002')  then // pt7
      begin
        // NATEXIS (5 car avec . décimal)
        SepDec := DecimalSeparator;
        DecimalSeparator := '.';
        if (length(FormatFloat('00.00',TobCdeFille.GetValue('PRT_VFACIALE'))) = 5) then
          VFaciale := FormatFloat('00.00',TobCdeFille.GetValue('PRT_VFACIALE'))
        else
          VFaciale := copy(HuitZero,1,(5-length(FormatFloat('00.00',TobCdeFille.GetValue('PRT_VFACIALE')))))+
                    FormatFloat('00.00',TobCdeFille.GetValue('PRT_VFACIALE'));
        DecimalSeparator := SepDec;
     // pt7 end;
     // deb pt7
     end
     else
     if (fournisseur = '004') then
     begin
      // CHEQUE DEJEUNER  (en cts sur 4 car)
       if (length(FormatFloat('0',TobCdeFille.GetValue('PRT_VFACIALE')*100)) = 4) then
          VFaciale := FormatFloat('0',TobCdeFille.GetValue('PRT_VFACIALE')*100)
        else
          VFaciale := copy(huitZero,1,(4-length(FormatFloat('0',TobCdeFille.GetValue('PRT_VFACIALE')*100))))+
                    FormatFloat('0',TobCdeFille.GetValue('PRT_VFACIALE')*100);
     end;
     // fin pt7
// f PT1-2

//d PT1-2
      if (Fournisseur = '001') or (Fournisseur = '003') then
      begin
        // SODEXHO  ou ACCOR (en cts sur 5 car)
        if (length(FormatFloat('0',TobCdeFille.GetValue('PRT_PPATRONALE')*100)) = 5) then
          PPatronale := FormatFloat('0',TobCdeFille.GetValue('PRT_PPATRONALE')*100)
        else
          PPatronale := copy(HuitZero,1,(5-length(FormatFloat('0',TobCdeFille.GetValue('PRT_PPATRONALE')*100))))+
                      FormatFloat('0',TobCdeFille.GetValue('PRT_PPATRONALE')*100);
      end
      else
      if (fournisseur = '002') then // pt7
      begin
       // NATEXIS (5 car avec . décimal)
       SepDec := DecimalSeparator;
       DecimalSeparator := '.';
       if (length(FormatFloat('00.00',TobCdeFille.GetValue('PRT_PPATRONALE'))) = 5) then
         PPatronale := FormatFloat('00.00',TobCdeFille.GetValue('PRT_PPATRONALE'))
       else
         PPatronale := copy(HuitZero,1,(5-length(FormatFloat('00.00',TobCdeFille.GetValue('PRT_PPATRONALE')))))+
                      FormatFloat('00.00',TobCdeFille.GetValue('PRT_PPATRONALE'));
        DecimalSeparator := SepDec;
    // pt7  end;
    // deb pt7
      end
      else
      if (fournisseur = '004') then
      begin
      // CHEQUE DEJEUNER (en cts sur 4 car)
        if (length(FormatFloat('0',TobCdeFille.GetValue('PRT_PPATRONALE')*100)) = 4) then
          PPatronale := FormatFloat('0',TobCdeFille.GetValue('PRT_PPATRONALE')*100)
        else
          PPatronale := copy(huitZero,1,(4-length(FormatFloat('0',TobCdeFille.GetValue('PRT_PPATRONALE')*100))))+
                      FormatFloat('0',TobCdeFille.GetValue('PRT_PPATRONALE')*100);
      end;
    // fin pt7
//f PT1-2

//d PT1-2
      if (Fournisseur = '001') then
      begin
      // SODEXHO
        if (length(TobCdeFille.GetValue('PRT_DISTRIBUTION')) >= 10) then
          Distribution := Copy(PgUpperCase(TobCdeFille.GetValue('PRT_DISTRIBUTION')),1,10)
        else
          Distribution := PgUpperCase(TobCdeFille.GetValue('PRT_DISTRIBUTION'))+
                          copy(blanc,1,(10-length(TobCdeFille.GetValue('PRT_DISTRIBUTION'))));
      end
      else
      if (Fournisseur = '002') then // PT6
      begin
      // NATEXIS
        if (length(TobCdeFille.GetValue('PRT_DISTRIBUTION')) >= 7) then
          Distribution := Copy(PgUpperCase(TobCdeFille.GetValue('PRT_DISTRIBUTION')),1,7)
        else
          Distribution := PgUpperCase(TobCdeFille.GetValue('PRT_DISTRIBUTION'))+
                          copy(blanc,1,(7-length(TobCdeFille.GetValue('PRT_DISTRIBUTION'))));
// d PT6
      end
      else
      if (Fournisseur = '003') then
      begin
      // ACCOR
          Distribution := PgUpperCase(TobCdeFille.GetValue('PRT_DISTRIBUTION'));
      // pt7 end;
      // deb pt7
      end
      else
      if (fournisseur =  '004') then
      begin
      // CHEQUE DEJEUNER
        abregedistrib := TobCdeFille.GetValue('PRT_DISTRIBUTION');
        QQ:= OpenSql('SELECT CC_CODE FROM CHOIXCOD '+
              'WHERE CC_ABREGE="'+abregedistrib+'"',True);

        if not QQ.EOF then distribution := QQ.Findfield('CC_CODE').asstring
        else distribution := '000';
        ferme(QQ);
      end;
      // fin pt7

// f PT1-2
    // pt7   if (Fournisseur = '001') or (Fournisseur = '002')then
     if (Fournisseur = '001') or (Fournisseur = '002') or (fournisseur = '004') then   // pt7
      // NATEXIS ou CHEQUE DEJEUNER
      begin
      if (length(TobCdeFille.GetValue('PRT_SALARIE')) = 11) then
        Salarie := PgUpperCase(TobCdeFille.GetValue('PRT_SALARIE'))
      else
        Salarie := PgUpperCase(TobCdeFille.GetValue('PRT_SALARIE'))+
                    copy(blanc,1,(11-length(TobCdeFille.GetValue('PRT_SALARIE'))));
      end
      else
      // ACCOR
      begin
      if (length(TobCdeFille.GetValue('PRT_SALARIE')) = 10) then
        Salarie := PgUpperCase(TobCdeFille.GetValue('PRT_SALARIE'))
      else
        Salarie := PgUpperCase(TobCdeFille.GetValue('PRT_SALARIE'))+
                    copy(blanc,1,(10-length(TobCdeFille.GetValue('PRT_SALARIE'))));
      end;

// f PT6
//d PT1-2
     // pt7 if (Fournisseur = '001') then
     if ((Fournisseur = '001') or (fournisseur = '004')) then     // pt7
      begin
      // SODEXHO  ou CHEQUE DEJEUNER
       if (length(TobCdeFille.GetValue('PRT_NOMPRENOM')) = 20) then
         NomPrenom := PgUpperCase(TobCdeFille.GetValue('PRT_NOMPRENOM'))
       else
         NomPrenom := PgUpperCase(TobCdeFille.GetValue('PRT_NOMPRENOM')) +
                      copy(blanc,1,(20-length(TobCdeFille.GetValue('PRT_NOMPRENOM'))));
      end
      else
      if (Fournisseur = '002') then
      begin
      // NATEXIS
       if (length(TobCdeFille.GetValue('PRT_NOMPRENOM')) = 32) then
         NomPrenom := PgUpperCase(TobCdeFille.GetValue('PRT_NOMPRENOM'))
       else
         NomPrenom := PgUpperCase(TobCdeFille.GetValue('PRT_NOMPRENOM')) +
                      copy(blanc,1,(32-length(TobCdeFille.GetValue('PRT_NOMPRENOM'))));
      end
      else
      if (Fournisseur = '003') then
      begin
      //ACCOR
       if (length(TobCdeFille.GetValue('PRT_NOMPRENOM')) = 40) then
         NomPrenom := PgUpperCase(TobCdeFille.GetValue('PRT_NOMPRENOM'))
       else
         NomPrenom := PgUpperCase(TobCdeFille.GetValue('PRT_NOMPRENOM')) +
                      copy(blanc,1,(40-length(TobCdeFille.GetValue('PRT_NOMPRENOM'))));
      end;
//f PT1-2

//d PT1-2
      if (Fournisseur = '001') then
      begin
        // SODEXHO
        if (length(TobCdeFille.GetValue('PRT_PERSONNAL')) = 14) then
          Personnal := PgUpperCase(TobCdeFille.GetValue('PRT_PERSONNAL'))
        else
          Personnal := PgUpperCase(TobCdeFille.GetValue('PRT_PERSONNAL'))+
                      copy(blanc,1,(14-length(TobCdeFille.GetValue('PRT_PERSONNAL'))));
      end
      else
      begin
      // NATEXIS
        if (length(TobCdeFille.GetValue('PRT_PERSONNAL')) = 32) then
          Personnal := PgUpperCase(TobCdeFille.GetValue('PRT_PERSONNAL'))
        else
          Personnal := PgUpperCase(TobCdeFille.GetValue('PRT_PERSONNAL'))+
                      copy(blanc,1,(32-length(TobCdeFille.GetValue('PRT_PERSONNAL'))));
      end;
//f PT1-2

    end
    else
    begin
      if (TobCdeFille  <> Nil) then
        SalPrec := TobCdeFille.GetValue('PRT_SALARIE');
// d PT4
      if(TobCdeFille <> Nil) and ( TobCdeFille.GetValue('PRT_NBTICKET') = 0)    then
//      if (NbT = 0) then
      begin
        Trace.Items.Add('!! Nbre de tickets = 0 : Matricule '+
                          SalPrec+
// f PT4
                          ' Code distribution '+
//                          Distribution);
                          TobCdeFille.GetValue('PRT_DISTRIBUTION'));
        Trace.Items.Add('');
// d PT2
        TobCdeFille := TobCde.FindNext ([''],[''], True);
//PT5       if (TobCdeFille = nil) then SalPrec := '';
// f PT2
        continue;
      end;
      if (NbT <> 0) then   // PT6
      begin                // PT6
        {Cumuls Nbre de tickets et Nbre de carnets}

        TotTicket := TotTicket + NbT;
        if fournisseur <> '004' then    // pt7
        begin                           // pt7
          if ((NbT div 25) * 25) < NbT then
            NbCarnet := NbCarnet + (NbT div 25)+1
          else
            NbCarnet := NbCarnet + (NbT div 25);
        end;                           // pt7

// d PT1-2
        if (Fournisseur = '001') then
        begin
        // SODEXHO

          NbLigne := (NbT div 99);
          if ((Nbligne * 99) < NbT) then Nbligne := NbLigne + 1;
          NbTicket := NbT;
          if (NbLigne > 1) then NbTicket := 99;
          while (NoOrdre < Nbligne) do
          begin
            Writeln(FPreFichier,WCodeClient+
                          Distribution+
                          ColleZeroDevant(NbTicket,2)+
                          VFaciale+
                          Salarie+
                          NomPrenom+
                          Personnal+
                          PPatronale+
                          '    '+
                          '2');

            {Traitement les lignes > 99}
            if (NbTicket = 99) then
              NbTicket := NbT - (NbTicket * (NoOrdre + 1));
            if (NbTicket > 99) then NbTicket := 99;
            NoOrdre :=  NoOrdre + 1;
          end;
        end
        else
        if (Fournisseur = '002') then // PT6
        begin
        // NATEXIS
          if ((NbT div 25) * 25) < NbT then
            NbCarnetLigne := (NbT div 25)+1
          else
            NbCarnetLigne := (NbT div 25);

          if (NbCarnetLigne > 999) then
          // Le nbre de carnets par ligne de cde est limité à 999
          begin
            NbCarnetW := 999;
            NbCarnetTotal := NbCarnetLigne;
            while (NbCarnetW < NbCarnetTotal) do
            begin
              NbcarnetLigne := 999;
              NbTicketCarnet := 25;
              NbTotalLigne := NbCarnetLigne * NbTicketCarnet;
              Writeln(FPreFichier,Distribution+
                              WCodeClient+
                              ColleZeroDevant(NbCarnetLigne,3)+
                              ColleZeroDevant(NbTicketCarnet,2)+
                              ColleZeroDevant(NbTotalLigne,5)+
                              VFaciale+
                              PPatronale+
                              NomPrenom+
                              Personnal+
                              '2');
              NbCarnetW := NbCarnetW + 999;
            end;
            NbCarnetW := NbCarnetW - 999;
            Reste := NbT - (NbCarnetW*25);
            NbT := Reste ;
            if ((NbT div 25) * 25) < NbT then
              NbCarnetLigne := (NbT div 25)+1
            else
              NbCarnetLigne := (NbT div 25);
          end;

          if ((NbCarnetLigne * 25) > NbT) and (NbT > 25) then
          begin
            NbCarnetLigne := NbCarnetLigne - 1;
            NbTicketCarnet := 25;
            NbTotalLigne := NbCarnetLigne * NbTicketCarnet;
            Writeln(FPreFichier,Distribution+
                            WCodeClient+
                            ColleZeroDevant(NbCarnetLigne,3)+
                            ColleZeroDevant(NbTicketCarnet,2)+
                            ColleZeroDevant(NbTotalLigne,5)+
                            VFaciale+
                            PPatronale+
                            NomPrenom+
                            Personnal+
                              '2');
            NbCarnetLigne :=  1;
            NbTicketCarnet := NbT - NbTotalLigne;
            NbTotalLigne := NbTicketCarnet;
            Writeln(FPreFichier,Distribution+
                                WCodeClient+
                                ColleZeroDevant(NbCarnetLigne,3)+
                                ColleZeroDevant(NbTicketCarnet,2)+
                                ColleZeroDevant(NbTotalLigne,5)+
                                VFaciale+
                                PPatronale+
                                NomPrenom+
                                Personnal+
                              '2');

          end
          else
          begin
            if (NbCarnetLigne <> 0) then      //PT5
              NbTicketCarnet  := trunc(NbT / NbCarnetLigne);
            NbTotalLigne :=  NbTicketCarnet * NbCarnetLigne;
            Writeln(FPreFichier,Distribution+
                              WCodeClient+
                              ColleZeroDevant(NbCarnetLigne,3)+
                              ColleZeroDevant(NbTicketCarnet,2)+
                              ColleZeroDevant(NbTotalLigne,5)+
                              VFaciale+
                              PPatronale+
                              NomPrenom+
                              Personnal+
                              '2');
          end;
// d PT6
//    end;
        end
        else
        if (Fournisseur = '003') then
        // ACCOR
        begin
          if (V_PGI.NumVersionBase < 800) then
          // (spécif V7)
          begin
            QQ := opensql('SELECT * FROM DEPORTSAL WHERE PSE_SALARIE = "'+Salarie+'"',TRUE);
            if not QQ.EOF then
            begin
              PointdeLivraison :=  QQ.findfield('PSE_MSAUNITEGES').asstring;
            end;
            ferme (QQ);
          end
          else
          // (à partir V8)
          begin
            QQ := opensql('SELECT * FROM DEPORTSAL WHERE PSE_SALARIE = "'+Salarie+'"',TRUE);
            if not QQ.EOF then
            begin
              PointdeLivraison :=  QQ.findfield('PSE_TICKLIVR').asstring;
            end;
            ferme (QQ);
          end;

          QQ := opensql('SELECT * FROM CHOIXCOD WHERE CC_TYPE="PCD" AND CC_ABREGE="'+Distribution+'"',TRUE);
          if not QQ.EOF then
          begin
            NiveauInterne1 :=  QQ.findfield('CC_LIBRE').asstring;
          end;
          ferme (QQ);
          if (length(NiveauInterne1) < 6) then
            NiveauInterne1 := NiveauInterne1 + copy(blanc,1,(6-length(NiveauInterne1)));

          NiveauInterne2 := '      ';

          for i:= 1 to 9 do
          begin
            if (Salarie[i] = '0')  then
              Salarie[i] := ' '
            else
              break;
          end;

          for i:= 1 to 5 do
          begin
            if (VFaciale[i] = '0')  then
              VFaciale[i] := ' '
            else
              break;
          end;
          if (VFaciale[5] = ' ') then
            VFaciale[5] := '0';

          for i:= 1 to 5 do
          begin
            if (PPatronale[i] = '0') then
              PPatronale[i] := ' '
            else
              break;
          end;
          if (PPatronale[5] = ' ') then
            PPatronale[5] := '0';

          ANbt := ColleZeroDevant(NbT,6);
          for i:= 1 to 6 do
          begin
            if (ANbt[i] = '0') then
              ANbt[i] := ' '
            else
              break;
          end;
          if (ANbt[6] = ' ') then
            ANbt[6] := '0';

          if (GetParamSocSecur ('SO_PGTKEDNPCARNET',FALSE) = TRUE) then
            EdNPCarnet := 'O'
          else
            EdNPCarnet := 'N';

          if(GetParamSocSecur ('SO_PGPERSOTICKET',FALSE) = TRUE) then
            EdNPTicket := 'O'
          else
            EdNPTicket := 'N';

          if (GetParamSocSecur ('SO_PGTKEDRSCARNET',FALSE) = TRUE) then
            EdRSCarnet := 'O'
          else
            EdRSCarnet := 'N';

          if (GetParamSocSecur ('SO_PGTKEDCPVCARNET',FALSE) = TRUE) then
            EdCPVCarnet := 'O'
          else
            EdCPVCarnet := 'N';

          CodePostal := GetParamSoc('SO_CODEPOSTAL');

          if (length(CodePostal) >= 5) then
            CodePostal :=  copy(CodePostal,1,5)
          else
            CodePostal := CodePostal + copy(blanc,1,(5-length(CodePostal)));

          Ville := GetParamSoc('SO_VILLE');
          if (length(Ville) >= 25) then
            Ville := copy(Ville,1,25)
          else
            Ville := Ville + copy(blanc,1,(25-length(Ville)));

          if (GetParamSocSecur ('SO_PGTKDATELIVR', FALSE) = TRUE) then
          // enregistrement avec date de livraison
          begin
            DateLivraison := GetControlText('DATELIVR');
            Writeln(FPreFichier,'50'+
                              ';'+
                              WCodeClient+
                              ';'+
                              PointdeLivraison+
                              ';'+
                              NiveauInterne1+               // point de distribution
                              ';'+
                              NiveauInterne2+
                              ';'+
                              Salarie+
                              ';'+
                              NomPrenom+
                              ';'+
                              EdNPCarnet+
                              ';'+
                              EdNPTicket+
                              ';'+
                              VFaciale+
                              ';'+
                              PPatronale+
                              ';'+
                              ANbt+
                              ';'+
                              Copy(WRaisonSociale,1,34)+
                              ';'+
                              CodePostal+
                              ';'+
                              Ville+
                              ';'+
                              EdRSCarnet+
                              ';'+
                              EdCPVCarnet+
                              ';'+
                              DateLivraison);
          end
          // enregistrement sans date de livraison
          else
            Writeln(FPreFichier,'50'+
                              ';'+
                              WCodeClient+
                              ';'+
                              PointdeLivraison+
                              ';'+
                              NiveauInterne1+               // point de distribution
                              ';'+
                              NiveauInterne2+
                              ';'+
                              Salarie+
                              ';'+
                              NomPrenom+
                              ';'+
                              EdNPCarnet+
                              ';'+
                              EdNPTicket+
                              ';'+
                              VFaciale+
                              ';'+
                              PPatronale+
                              ';'+
                              ANbt+
                              ';'+
                              Copy(WRaisonSociale,1,34)+
                              ';'+
                              CodePostal+
                              ';'+
                              Ville+
                              ';'+
                              EdRSCarnet+
                              ';'+
                              EdCPVCarnet);

    // pt7   end;//PT6
    // deb pt7
       end
       else
       if (fournisseur = '004') then
       begin
        if (NBT >= 1) then
        begin

         // si nb tickets > 50 alors génére plusieurs lignes dans le fichier
         if NBT <= 50 then nbcarnetligne := 1
         else
          if (((NBT div 50) * 50) = NBT) then nbcarnetligne := (NBT div 50) else
          nbcarnetligne := (NBT div 50) + 1;
          nbcarnet := nbcarnet + nbcarnetligne;
          
          nbticketprec := 0;
           For i := 1 to nbcarnetligne do
           begin
            if NBT <= 50 then nbticketcarnet := NBT else
            nbticketcarnet := 50;
            if (nbticketcarnet + nbticketprec) > NBT then nbticketcarnet := NBT - nbticketprec;
            if (nbticketcarnet > 0) and (nbticketcarnet <= 50) then
            begin
             rang := rang + 1;

             writeln(FPreFichier, '4' +  '1' +
             wcodeclient +
             distribution +
             huitblanc +
             salarie +
             Collezerodevant(rang,6) +
             Nomprenom +
             collezerodevant(nbticketcarnet, 2) +
             '0' +
             Vfaciale +
             Ppatronale +
             '00');
            end;
          nbticketprec := nbticketprec + nbticketcarnet;
         end;
        end;
       end;
        // fin pt7
    end;
// f PT6
// f PT1-2
      if (TobCdeFille = Nil) then SalPrec := '';
      NbT := 0;
      continue;
    end;

    TobCdeFille := TobCde.FindNext ([''],[''], True);
  end;
  if (CodeRetour = False) then exit;

  {Enregitrement de fin de fichier}

    writeln(FPreFichier,'********'+
                      ColleZeroDevant(TotTicket,6)+
                      ColleZeroDevant(NbCarnet,6)+
                       copy(blanc,1,59)+
                       '3');
    Trace.Items.Add ('Nombre de carnets commandés  : '+
                   IntToStr(NbCarnet));
    Trace.Items.Add ('Nombre de titres commandés   : '+
                   IntToStr(TotTicket));

end;

Initialization
  registerclasses ( [ TOF_PGPREPCDETK ] ) ;
end.
