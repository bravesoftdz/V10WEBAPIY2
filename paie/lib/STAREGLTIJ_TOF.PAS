{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 14/09/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : STAREGLTIJ ()
Mots clefs ... : TOF;STAREGLTIJ
*****************************************************************}
{
 PT1    : 12/01/2005 MF V_6.0 :   Correction FQ 11883 + FQ 11884 + FQ 11885
                                  + possibilité de sélectioner Toutes les
                                  absences ou uniquement celles réglées ou
                                  uniquement celles non réglées
 PT2    : 28/01/2005 MF V_6.0 :   Modification pour avoir accès aux champs de
                                  ABSENCESALARIE dans l'onglet avancé
 PT3    : 08/02/2005 MF V_6.0 :   Ajout Date d'ancienneté ds la liste.
 PT4    : 21/02/2005 MF V_6.0 :   Formatage matricule salarié
 PT5    : 22/02/2005 MF v_6.0 :   correction erreur "indice de liste hors limite"
 PT6    : 17/06/2005 MF V_6.02 :  Modification des requêtes afin de récupérer les
                                  absence dont la date de fin est comprise dans
                                  la période sélectionnée.
PT7     : 23/01/2006 SB V_65 FQ 10866 Ajout clause predefini motif d'absence
PT8     : 15/02/2006 MF V_65 FQ 12712 correction requête
PT9     : 09/03/2006 MF V_65 Correction alimentation date à Idate1900. Corrige erreur
                             sur liste d'export en CWAS (30/12/1899)
PT10   19/06/2006 SB V_65 FQ 13231 Retrait des mvt absences annulées
}
Unit STAREGLTIJ_TOF ;

Interface

Uses StdCtrls,
//unused     Controls,
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
//unused     mul,
{$else}
//unused     eMul, 
{$ENDIF}
     EntPaie,
//unused     forms,
     sysutils, 
     ComCtrls,
     HCtrls,
     HEnt1, 
//unused     HMsgBox,
     HQry,
     PgOutils2,
     stat,
     uTob,
     UTOF ;

Type
  TOF_STAREGLTIJ = Class (TOF)
  private
  T_Reglt,T_Absence                       : TOB;
// d PT2
  Stwhere                                 : string;
  BoolStwhere                             : boolean;
// f PT2

  procedure LanceExtract(var XX_WHERE, Salarie, Etablissement : string;ChkRegle : TCheckBox); // PT1
  procedure ConstructionLaTob(const DateRegltON, NonRegle : boolean); // PT1
  procedure ExitEdit(Sender: TObject);
  procedure ChangeChk(Sender: TObject);
  procedure ExitEditS(Sender: TObject);     // PT4


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

procedure TOF_STAREGLTIJ.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_STAREGLTIJ.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_STAREGLTIJ.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_STAREGLTIJ.OnLoad ;
var
  DateDebAbs, DateFinAbs, DateDebReglt,DateFinReglt  : string;    // PT1
  XX_WHERE                                           : string;
  TypeAbs,Salarie , Etablissement                    : string;
  DateRegltON, NonRegle                              : boolean;   // PT1
  ChkRegle,ChkRegle1                                 : TCheckBox; // PT1
//  stWhere                                            : string;
begin
  Inherited ;
  if Assigned(T_Reglt)  then FreeAndNil (T_Reglt);
  if Assigned(T_Absence)  then FreeAndNil (T_Absence);

   XX_WHERE := '';

   DateDebAbs := (THEdit(Getcontrol('DATEDEBUTABS')).Text);
   XX_WHERE := DateDebAbs;

   DateFinAbs := (THEdit(Getcontrol('DATEFINABS')).Text);
   XX_WHERE := XX_WHERE +';'+DateFinAbs;

   DateDebReglt := (THEdit(Getcontrol('DATEDEBREGLT')).Text);    // PT1
   XX_WHERE := XX_WHERE +';'+ DateDebReglt;

   DateFinReglt := (THEdit(Getcontrol('DATEFINREGLT')).Text);    // PT1
   XX_WHERE := XX_WHERE +';'+ DateFinReglt;


   TypeAbs := (THValComboBox(GetControl('TYPEABS')).value);
   XX_WHERE := XX_WHERE +';'+ TypeAbs ;

   Salarie := (THMultiValComboBox(GetControl('SALARIE')).value);
   Etablissement := (THMultiValComboBox(GetControl('ETABLISSEMENT')).value);

   DateRegltON := false;
   if (THEdit(GetControl('DATEDEBREGLT')).Text <> '  /  /    ') and
      (THEdit(GetControl('DATEFINREGLT')).Text <> '  /  /    ') then    // PT1
     DateRegltON := true;


// d PT1
   NonRegle := true;
   ChkRegle := TCheckBox(GetControl('CHKREGLE'));
   ChkRegle1 := TCheckBox(GetControl('CHKREGLE1'));
   if (ChkRegle.checked=False) then
     DateRegltON := false;    // toutes les absences
   if (ChkRegle1.checked=False) then
     NonRegle := false;       // uniqt absences non réglées
// f PT1
   LanceExtract(XX_WHERE,Salarie,Etablissement,ChkRegle);
   ConstructionLaTOB(DateRegltON,NonRegle); //PT1
// f PT1
end ;

procedure TOF_STAREGLTIJ.OnArgument (S : String ) ;
var
  Defaut                               : THEdit;
  Defaut1                              : THEdit; // PT4
  ChkRegle,ChkRegle1                   : TCheckBox; //PT1
begin
  Inherited ;
// d PT1
  Defaut:=ThEdit(getcontrol('DATEDEBREGLT'));
  if Defaut<>nil then
  begin
    Defaut.OnExit:=ExitEdit;
    Defaut.OnChange := ExitEdit;
  end;
  ChkRegle := TCheckBox(getcontrol('CHKREGLE'));
  if (ChkRegle <> nil) then
  begin
    ChkRegle.Onexit := Changechk;
  end;
  ChkRegle1 := TCheckBox(getcontrol('CHKREGLE1'));
  if (ChkRegle1 <> nil) then
  begin
    ChkRegle1.Onexit := Changechk;
  end;
// f PT1
// d PT1
  Defaut1:=ThEdit(getcontrol('SALARIE'));
  if Defaut1<>nil then Defaut1.OnExit:=ExitEditS;
// f PT1

end ;

procedure TOF_STAREGLTIJ.OnClose ;
begin
  Inherited ;
  FreeAndNil (T_Reglt);
  FreeAndNil (T_Absence);
end ;

procedure TOF_STAREGLTIJ.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_STAREGLTIJ.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_STAREGLTIJ.LanceExtract(var XX_WHERE, Salarie, Etablissement : string;ChkRegle : TCheckBox);
var
   Q                                                           : TQuery;
   RWhereTmp,RWhere, StCondition                               : string;
   AWhereTmp,AWhere                                            : string; 
   DateDebAbs, DateFinAbs, DateDebReglt,DateFinREglt,TypeAbs   : string;  //PT1
   Criteres                                                    : boolean;
   i                                                           : integer;

begin
// d PT2
  // récupération du Where de la requête de la fiche
  // BooStWhre = true quand des critères ont été saisis sur l'onglet avancé
  BoolStWhere := false;
  StWhere := RecupWhereCritere(TPageControl(GetControl('Pages')));
  if (StWhere = '') then
    StWhere :=' WHERE '
  else
    BoolStWhere := true;
// f PT2

  criteres := false;

  if (XX_WHERE <> '') or (Salarie <> '') or (Etablissement <> '') then
    Criteres := true;

  DateDebAbs := ReadTokenPipe(XX_WHERE,';');
  DateFinAbs := ReadTokenPipe(XX_WHERE,';');
  DateDebReglt  := ReadTokenPipe(XX_WHERE,';');     //PT1
  DateFinReglt  := ReadTokenPipe(XX_WHERE,';');     //PT1
  TypeAbs    := ReadTokenPipe(XX_WHERE,';');

  // Chargement des tobs
  // ===================

  // construction requêtes
  if Criteres then
  begin
    AWhereTmp := '';
    RWhereTmp := '';

    Rwhere := 'WHERE ';

    While (Etablissement <> '') do
    begin
      StCondition := ReadTokenPipe(Etablissement,';');
      If RWhereTmp <> '' then
      begin
        RWhereTmp := RWhereTmp  + ' OR PSA_ETABLISSEMENT="'+StCondition+'"';
        AWhereTmp := AWhereTmp  + ' OR PCN_ETABLISSEMENT="'+StCondition+'"';
      end
      else
        if (Rwhere = 'WHERE ') then
        begin
          RWhereTmp := '(PSA_ETABLISSEMENT="'+StCondition+'"';
          AWhereTmp := '(PCN_ETABLISSEMENT="'+StCondition+'"';
        end
        else
        begin
          RWhereTmp := ' AND (PSA_ETABLISSEMENT="'+StCondition+'"';
          AWhereTmp := ' AND (PCN_ETABLISSEMENT="'+StCondition+'"';
        end;
    end;
    if RWhereTmp <> '' then
      RWhereTmp := RWhereTmp + ')';
   if AWhereTmp <> '' then
      AWhereTmp := AWhereTmp + ')';

      Rwhere := Rwhere + RWhereTmp;
      Awhere := Awhere + AWhereTmp;

    RWhereTmp := '';
    AWhereTmp := '';
    While Salarie <> '' do
    begin
      StCondition := ReadTokenPipe(Salarie,';');
      If RWhereTmp <> '' then
      begin
        RWhereTmp := RWhereTmp  + ' OR PRI_SALARIE="'+StCondition+'"';
        AWhereTmp := AWhereTmp  + ' OR PCN_SALARIE="'+StCondition+'"';
      end
      else
        if (Rwhere = 'WHERE ') then
        begin
          RWhereTmp := '(PRI_SALARIE="'+StCondition+'"';
          AWhereTmp := '(PCN_SALARIE="'+StCondition+'"';
        end
        else
        begin
          RWhereTmp := ' AND (PRI_SALARIE="'+StCondition+'"';
          AWhereTmp := ' AND (PCN_SALARIE="'+StCondition+'"';
        end;
    end;
    if RWhereTmp <> '' then
       RWhereTmp := RWhereTmp + ')';
    if AWhereTmp <> '' then
      AWhereTmp := AWhereTmp + ')';

      Rwhere := Rwhere + RWhereTmp;
      Awhere := Awhere + AWhereTmp;
  end;

  if (DateDebAbs <> '') and (DateDebAbs <> '  /  /    ') then
    if (Rwhere = 'WHERE ') then
    begin
//PT6      Rwhere := Rwhere + 'PRI_DATEDEBUTABS >="'+UsDateTime(StrToDate(DateDebAbs))+'"';
      Rwhere := Rwhere + 'PRI_DATEFINABS >="'+UsDateTime(StrToDate(DateDebAbs))+'"';
//PT6      Awhere := Awhere + 'PCN_DATEDEBUTABS >="'+UsDateTime(StrToDate(DateDebAbs))+'"';
      Awhere := Awhere + 'PCN_DATEFINABS >="'+UsDateTime(StrToDate(DateDebAbs))+'"';
    end
    else
    begin
//PT6      Rwhere := Rwhere + ' AND PRI_DATEDEBUTABS >="'+UsDateTime(StrToDate(DateDebAbs))+'"';
      Rwhere := Rwhere + ' AND PRI_DATEFINABS >="'+UsDateTime(StrToDate(DateDebAbs))+'"';
//PT6      Awhere := Awhere + ' AND PCN_DATEDEBUTABS >="'+UsDateTime(StrToDate(DateDebAbs))+'"';
      Awhere := Awhere + ' AND PCN_DATEFINABS >="'+UsDateTime(StrToDate(DateDebAbs))+'"';
    end;
  if (DateFinAbs <> '') and (DateFinAbs <> '  /  /    ') then
    if (Rwhere = 'WHERE ') then
    begin
      Rwhere := Rwhere + 'PRI_DATEFINABS <="'+UsDateTime(StrToDate(DateFinAbs))+'"';
      Awhere := Awhere + 'PCN_DATEFINABS <="'+UsDateTime(StrToDate(DateFinAbs))+'"';
    end
    else
    begin
      Rwhere := Rwhere + ' AND PRI_DATEFINABS <="'+UsDateTime(StrToDate(DateFinAbs))+'"';
      Awhere := Awhere + ' AND PCN_DATEFINABS <="'+UsDateTime(StrToDate(DateFinAbs))+'"';
    end;

// d PT1
  if (ChkRegle <> nil) and (ChkRegle.Checked=True) then
  // on récupère toutes les absences sauf celles non réglées
  begin
    if (DateDebReglt <> '') and (DateDebReglt <> '  /  /    ') then
      if (Rwhere = 'WHERE ') then
        Rwhere := Rwhere + 'PRI_DATEREGLT >="'+UsDateTime(StrToDate(DateDebReglt))+'"'
      else
        Rwhere := Rwhere + ' AND PRI_DATEREGLT >="'+UsDateTime(StrToDate(DateDebReglt))+'"';

    if (DateFinReglt <> '') and (DateFinReglt <> '  /  /    ') then
      if (Rwhere = 'WHERE ') then
        Rwhere := Rwhere + 'PRI_DATEREGLT <="'+UsDateTime(StrToDate(DateFinReglt))+'"'
      else
        Rwhere := Rwhere + ' AND PRI_DATEREGLT <="'+UsDateTime(StrToDate(DateFinReglt))+'"';
  end;

// f PT1
// d PT2
  // alimentation requête sur ABSENCESALARIE
// d PT6 correction requête "WHERE AND..."
//PT8  if (StWhere ='WHERE') then
  if (StWhere = ' WHERE ') then
    StWhere := StWhere + ' '+ Awhere
  else
// f PT6
    StWhere := StWhere + ' AND '+ Awhere;

  Awhere := StWhere;
// f PT2

  if (TypeAbs <> '') then
  begin
    // alimentation requête sur REGLTIJSS
    if (Rwhere = 'WHERE ') then
      Rwhere := Rwhere + 'PRI_TYPEABS ="'+TypeAbs+'"'
    else
      Rwhere := Rwhere + ' AND PRI_TYPEABS ="'+TypeAbs+'"';
// d PT8
    if (Awhere = 'WHERE ') then
        Awhere := Awhere + ' PCN_TYPECONGE IN (SELECT PMA_MOTIFABSENCE '+
                           'FROM MOTIFABSENCE WHERE ##PMA_PREDEFINI## AND '+
                           'PMA_GESTIONIJSS="X" '
    else
        Awhere := Awhere + ' AND PCN_TYPECONGE IN (SELECT PMA_MOTIFABSENCE '+
                           'FROM MOTIFABSENCE WHERE ##PMA_PREDEFINI## AND '+
                           'PMA_GESTIONIJSS="X" ';
// f PT8

    if (length(TypeAbs) = 0) then
      Awhere := Awhere + ' AND PMA_TYPEABS <>"")'
    else
    begin
      Awhere := Awhere + ' AND (';
      i := 1;
      while (i < Length(TypeAbs)) do
      begin
        Awhere := Awhere + 'PMA_TYPEABS = "'+ Copy(TypeAbs,i,3)+'"';
        if (i+3 < Length(TypeAbs)) then
            Awhere := Awhere + ' OR ';
        i := i + 4;
      end;
      Awhere := Awhere + ')';
    end;
    Awhere := Awhere + ')';
  end;

  if  (Rwhere = 'WHERE ') then Rwhere := '';

  T_Reglt := TOB.Create('MaTob-Reglt', nil, -1);
  Q := OpenSql('select REGLTIJSS.*,PSA_ETABLISSEMENT,PSA_LIBELLE,PSA_PRENOM '+  //PT1
               'from REGLTIJSS '+
               'LEFT JOIN SALARIES ON PSA_SALARIE=PRI_SALARIE '+Rwhere, TRUE);
  if not q.eof then
    T_Reglt.LoadDetailDB('REGLTIJSS', '', '', Q, FALSE, FALSE);
  Ferme(Q);

  if (AWhere = 'WHERE ') then
    AWhere := Awhere +'PCN_GESTIONIJSS="X" '
  else
    AWhere := AWhere + ' AND PCN_GESTIONIJSS="X" ';

  AWhere := Awhere + ' AND PCN_ETATPOSTPAIE <> "NAN" '; { PT10 }

  T_Absence := TOB.Create('MaTob-Absences', nil, -1);
  Q := OpenSql('select ABSENCESALARIE.*,PSA_LIBELLE,PSA_PRENOM,'+
               'PSA_DATEANCIENNETE from ABSENCESALARIE '+        // PT3
               'LEFT JOIN SALARIES ON PSA_SALARIE=PCN_SALARIE '+Awhere, TRUE);
  if not q.eof then
    T_Absence.LoadDetailDB('ABSENCESALARIE', '', '', Q, FALSE, FALSE);
  Ferme(Q);
end;
procedure TOF_STAREGLTIJ.ConstructionLaTob(const DateRegltON, NonRegle : boolean); // PT1
var
  TA,FusionFille,AbsenceFille : TOB;
  i, Wi                       : integer;
  Q                           : TQuery;
  NbReglt                     : integer; // PT5
begin

   for i:= 0 to T_Reglt.Detail.Count-1 do
   begin
     FusionFille := T_Reglt.Detail[i];
     if (i=0) then
     begin
       FusionFille.AddChampSup('LIBELLE',TRUE);
// d PT2
         FusionFille.AddChampSup('PCN_TYPEMVT',TRUE);
         FusionFille.AddChampSup('PCN_RESSOURCE',TRUE);
         FusionFille.AddChampSup('PCN_DATEDEBUT',TRUE);
         FusionFille.AddChampSup('PCN_DATEFIN',TRUE);
         FusionFille.AddChampSup('PCN_TYPECONGE',TRUE);
         FusionFille.AddChampSup('PCN_SENSABS',TRUE);
         FusionFille.AddChampSup('PCN_MVTORIGINE',TRUE);
         FusionFille.AddChampSup('PCN_TYPEIMPUTE',TRUE);
         FusionFille.AddChampSup('PCN_DATECREATION',TRUE);
         FusionFille.AddChampSup('PCN_LIBELLE',TRUE);
         FusionFille.AddChampSup('PCN_GENERECLOTURE',TRUE);
         FusionFille.AddChampSup('PCN_DATEMODIF',TRUE);
         FusionFille.AddChampSup('PCN_DATESOLDE',TRUE);
         FusionFille.AddChampSup('PCN_DATEVALIDITE',TRUE);
         FusionFille.AddChampSup('PCN_DATEDEBUTABS',TRUE);
         FusionFille.AddChampSup('PCN_DEBUTDJ',TRUE);
         FusionFille.AddChampSup('PCN_DATEFINABS',TRUE);
         FusionFille.AddChampSup('PCN_FINDJ',TRUE);
         FusionFille.AddChampSup('PCN_DATEPAIEMENT',TRUE);
         FusionFille.AddChampSup('PCN_JOURS',TRUE);
         FusionFille.AddChampSup('PCN_HEURES',TRUE);
         FusionFille.AddChampSup('PCN_BASE',TRUE);
         FusionFille.AddChampSup('PCN_NBREMOIS',TRUE);
         FusionFille.AddChampSup('PCN_CODERGRPT',TRUE);
         FusionFille.AddChampSup('PCN_MVTDUPLIQUE',TRUE);
         FusionFille.AddChampSup('PCN_ABSENCE',TRUE);
         FusionFille.AddChampSup('PCN_ABSENCEMANU',TRUE);
         FusionFille.AddChampSup('PCN_MODIFABSENCE',TRUE);
         FusionFille.AddChampSup('PCN_APAYES',TRUE);
         FusionFille.AddChampSup('PCN_VALOX',TRUE);
         FusionFille.AddChampSup('PCN_VALOMS',TRUE);
         FusionFille.AddChampSup('PCN_VALORETENUE',TRUE);
         FusionFille.AddChampSup('PCN_VALOMANUELLE',TRUE);
         FusionFille.AddChampSup('PCN_MODIFVALO',TRUE);
         FusionFille.AddChampSup('PCN_PERIODEPAIE',TRUE);
         FusionFille.AddChampSup('PCN_ETABLISSEMENT',TRUE);
         FusionFille.AddChampSup('PCN_TRAVAILN1',TRUE);
         FusionFille.AddChampSup('PCN_TRAVAILN2',TRUE);
         FusionFille.AddChampSup('PCN_TRAVAILN3',TRUE);
         FusionFille.AddChampSup('PCN_TRAVAILN4',TRUE);
         FusionFille.AddChampSup('PCN_CODESTAT',TRUE);
         FusionFille.AddChampSup('PCN_CONFIDENTIEL',TRUE);
         FusionFille.AddChampSup('PCN_SAISIEDEPORTEE',TRUE);
         FusionFille.AddChampSup('PCN_VALIDSALARIE',TRUE);
         FusionFille.AddChampSup('PCN_VALIDRESP',TRUE);
         FusionFille.AddChampSup('PCN_EXPORTOK',TRUE);
         FusionFille.AddChampSup('PCN_LIBCOMPL1',TRUE);
         FusionFille.AddChampSup('PCN_LIBCOMPL2',TRUE);
         FusionFille.AddChampSup('PCN_VALIDABSENCE',TRUE);
         FusionFille.AddChampSup('PCN_OKFRACTION',TRUE);
         FusionFille.AddChampSup('PSA_DATEANCIENNETE',TRUE); // PT3
// f PT2
       FusionFille.AddChampSup('PCN_NBJCARENCE',TRUE);
       FusionFille.AddChampSup('PCN_NBJCALEND',TRUE);
       FusionFille.AddChampSup('PCN_NBJIJSS',TRUE);
       FusionFille.AddChampSup('PCN_IJSSSOLDEE',TRUE);
       FusionFille.AddChampSup('PCN_DATEPAIEMENT',TRUE);
       FusionFille.AddChampSup('LIBTYPEABS',TRUE);
       FusionFille.AddChampSup('MATRICULENOMPRENOM',TRUE);
     end;
     FusionFille.PutValue('LIBELLE','');
// d PT2
       FusionFille.PutValue('PCN_TYPEMVT','');
       FusionFille.PutValue('PCN_RESSOURCE','');
       FusionFille.PutValue('PCN_DATEDEBUT',IDate1900);
         FusionFille.PutValue('PCN_DATEFIN',IDate1900);
         FusionFille.PutValue('PCN_TYPECONGE','');
         FusionFille.PutValue('PCN_SENSABS','');
         FusionFille.PutValue('PCN_MVTORIGINE','');
         FusionFille.PutValue('PCN_TYPEIMPUTE','');
         FusionFille.PutValue('PCN_DATECREATION',IDate1900);
         FusionFille.PutValue('PCN_LIBELLE','');
         FusionFille.PutValue('PCN_GENERECLOTURE','');
         FusionFille.PutValue('PCN_DATEMODIF',IDate1900);
         FusionFille.PutValue('PCN_DATESOLDE',IDate1900);
         FusionFille.PutValue('PCN_DATEVALIDITE',IDate1900);
         FusionFille.PutValue('PCN_DATEDEBUTABS',IDate1900);
         FusionFille.PutValue('PCN_DEBUTDJ','');
         FusionFille.PutValue('PCN_DATEFINABS',IDate1900);
         FusionFille.PutValue('PCN_FINDJ','');
         FusionFille.PutValue('PCN_DATEPAIEMENT',IDate1900);
         FusionFille.PutValue('PCN_JOURS',0);
         FusionFille.PutValue('PCN_HEURES',0);
         FusionFille.PutValue('PCN_BASE',0);
         FusionFille.PutValue('PCN_NBREMOIS',0);
         FusionFille.PutValue('PCN_CODERGRPT',0);
         FusionFille.PutValue('PCN_MVTDUPLIQUE','');
         FusionFille.PutValue('PCN_ABSENCE',0);
         FusionFille.PutValue('PCN_ABSENCEMANU',0);
         FusionFille.PutValue('PCN_MODIFABSENCE','');
         FusionFille.PutValue('PCN_APAYES',0);
         FusionFille.PutValue('PCN_VALOX',0);
         FusionFille.PutValue('PCN_VALOMS',0);
         FusionFille.PutValue('PCN_VALORETENUE',0);
         FusionFille.PutValue('PCN_VALOMANUELLE',0);
         FusionFille.PutValue('PCN_MODIFVALO','');
         FusionFille.PutValue('PCN_PERIODEPAIE','');
         FusionFille.PutValue('PCN_ETABLISSEMENT','');
         FusionFille.PutValue('PCN_TRAVAILN1','');
         FusionFille.PutValue('PCN_TRAVAILN2','');
         FusionFille.PutValue('PCN_TRAVAILN3','');
         FusionFille.PutValue('PCN_TRAVAILN4','');
         FusionFille.PutValue('PCN_CODESTAT','');
         FusionFille.PutValue('PCN_CONFIDENTIEL','');
         FusionFille.PutValue('PCN_SAISIEDEPORTEE','');
         FusionFille.PutValue('PCN_VALIDSALARIE','');
         FusionFille.PutValue('PCN_VALIDRESP','');
         FusionFille.PutValue('PCN_EXPORTOK','');
         FusionFille.PutValue('PCN_LIBCOMPL1','');
         FusionFille.PutValue('PCN_LIBCOMPL2','');
         FusionFille.PutValue('PCN_VALIDABSENCE','');
         FusionFille.PutValue('PCN_OKFRACTION','');
         FusionFille.PutValue('PSA_DATEANCIENNETE',IDate1900);  // PT3

// f PT2
     FusionFille.PutValue('PCN_NBJCARENCE',0);
     FusionFille.PutValue('PCN_NBJCALEND',0);
     FusionFille.PutValue('PCN_NBJIJSS',0);
     FusionFille.PutValue('PCN_IJSSSOLDEE','');
     FusionFille.PutValue('PCN_DATEPAIEMENT',IDate1900);   // PT9
     FusionFille.PutValue('LIBTYPEABS',
                          RechDom('PGTYPEABS',FusionFille.GetValue('PRI_TYPEABS'),FALSE));
     FusionFille.PutValue('MATRICULENOMPRENOM',
                          Trim(FusionFille.GetValue('PRI_SALARIE'))+
                          '  '+
                          Trim(FusionFille.GetValue('PSA_LIBELLE'))+
                          '  '+
                          Trim(FusionFille.GetValue('PSA_PRENOM')));
   end;
// d PT2
   // lorsque des critères ont été saisi sur l'onglet avancé il faut éliminer
   // les reglements non concernés.
// d PT5
   Wi := 0;
   NbReglt :=  T_Reglt.Detail.Count-1;
   for i := 0 to  NbReglt do
   begin
     if (BoolStWhere) then
     begin
       FusionFille := T_Reglt.Detail[Wi];
       AbsenceFille := T_Absence.FindFirst(['PCN_SALARIE','PCN_DATEDEBUTABS','PCN_DATEFINABS'],
                    [FusionFille.getValue('PRI_SALARIE'),FusionFille.getValue('PRI_DATEDEBUTABS'),
                     FusionFille.getValue('PRI_DATEFINABS')],
                    TRUE);
       if (AbsenceFille = nil) then
       begin
         FusionFille.free;
       end
       else
         wi := wi + 1 ;
     end;
// f PT5
// f PT2
   end;

   for i:= 0 to T_Absence.Detail.Count-1 do
   begin
     TA := T_Absence.Detail[i];

     FusionFille :=  T_Reglt.FindFirst (['PRI_SALARIE','PRI_DATEDEBUTABS','PRI_DATEFINABS'],
                    [TA.getValue('PCN_SALARIE'),TA.getValue('PCN_DATEDEBUTABS'),
                     TA.getValue('PCN_DATEFINABS')],
                    TRUE);
     if (FusionFille <> nil) then
     // un élt. déjà présent pour cette absence (un règlement)
     begin
       if (not NonRegle) then // PT1
       // un élt. déjà présent pour cette absence (un règlement)
       begin

         FusionFille.PutValue('LIBELLE', TA.getValue('PCN_LIBELLE')); // libellé absence
// PT2
         FusionFille.PutValue('PCN_DATEFIN',TA.getValue('PCN_DATEFIN'));
         FusionFille.PutValue('PCN_TYPECONGE',TA.getValue('PCN_TYPECONGE'));
         FusionFille.PutValue('PCN_SENSABS',TA.getValue('PCN_SENSABS'));
         FusionFille.PutValue('PCN_MVTORIGINE',TA.getValue('PCN_MVTORIGINE'));
         FusionFille.PutValue('PCN_TYPEIMPUTE',TA.getValue('PCN_TYPEIMPUTE'));
         FusionFille.PutValue('PCN_DATECREATION',TA.getValue('PCN_DATECREATION'));
         FusionFille.PutValue('PCN_LIBELLE',TA.getValue('PCN_LIBELLE'));
         FusionFille.PutValue('PCN_GENERECLOTURE',TA.getValue('PCN_GENERECLOTURE'));
         FusionFille.PutValue('PCN_DATEMODIF',TA.getValue('PCN_DATEMODIF'));
         FusionFille.PutValue('PCN_DATESOLDE',TA.getValue('PCN_DATESOLDE'));
         FusionFille.PutValue('PCN_DATEVALIDITE',TA.getValue('PCN_DATEVALIDITE'));
         FusionFille.PutValue('PCN_DATEDEBUTABS',TA.getValue('PCN_DATEDEBUTABS'));
         FusionFille.PutValue('PCN_DEBUTDJ',TA.getValue('PCN_DEBUTDJ'));
         FusionFille.PutValue('PCN_DATEFINABS',TA.getValue('PCN_DATEFINABS'));
         FusionFille.PutValue('PCN_FINDJ',TA.getValue('PCN_FINDJ'));
         FusionFille.PutValue('PCN_DATEPAIEMENT',TA.getValue('PCN_DATEPAIEMENT'));
         FusionFille.PutValue('PCN_JOURS',TA.getValue('PCN_JOURS'));
         FusionFille.PutValue('PCN_HEURES',TA.getValue('PCN_HEURES'));
         FusionFille.PutValue('PCN_BASE',TA.getValue('PCN_BASE'));
         FusionFille.PutValue('PCN_NBREMOIS',TA.getValue('PCN_NBREMOIS'));
         FusionFille.PutValue('PCN_CODERGRPT',TA.getValue('PCN_CODERGRPT'));
         FusionFille.PutValue('PCN_MVTDUPLIQUE',TA.getValue('PCN_MVTDUPLIQUE'));
         FusionFille.PutValue('PCN_ABSENCE',TA.getValue('PCN_ABSENCE'));
         FusionFille.PutValue('PCN_ABSENCEMANU',TA.getValue('PCN_ABSENCEMANU'));
         FusionFille.PutValue('PCN_MODIFABSENCE',TA.getValue('PCN_MODIFABSENCE'));
         FusionFille.PutValue('PCN_APAYES',TA.getValue('PCN_APAYES'));
         FusionFille.PutValue('PCN_VALOX',TA.getValue('PCN_VALOX'));
         FusionFille.PutValue('PCN_VALOMS',TA.getValue('PCN_VALOMS'));
         FusionFille.PutValue('PCN_VALORETENUE',TA.getValue('PCN_VALORETENUE'));
         FusionFille.PutValue('PCN_VALOMANUELLE',TA.getValue('PCN_VALOMANUELLE'));
         FusionFille.PutValue('PCN_MODIFVALO',TA.getValue('PCN_MODIFVALO'));
         FusionFille.PutValue('PCN_PERIODEPAIE',TA.getValue('PCN_PERIODEPAIE'));
         FusionFille.PutValue('PCN_ETABLISSEMENT',TA.getValue('PCN_ETABLISSEMENT'));
         FusionFille.PutValue('PCN_TRAVAILN1',TA.getValue('PCN_TRAVAILN1'));
         FusionFille.PutValue('PCN_TRAVAILN2',TA.getValue('PCN_TRAVAILN2'));
         FusionFille.PutValue('PCN_TRAVAILN3',TA.getValue('PCN_TRAVAILN3'));
         FusionFille.PutValue('PCN_TRAVAILN4',TA.getValue('PCN_TRAVAILN4'));
         FusionFille.PutValue('PCN_CODESTAT',TA.getValue('PCN_CODESTAT'));
         FusionFille.PutValue('PCN_CONFIDENTIEL',TA.getValue('PCN_CONFIDENTIEL'));
         FusionFille.PutValue('PCN_SAISIEDEPORTEE',TA.getValue('PCN_SAISIEDEPORTEE'));
         FusionFille.PutValue('PCN_VALIDSALARIE',TA.getValue('PCN_VALIDSALARIE'));
         FusionFille.PutValue('PCN_VALIDRESP',TA.getValue('PCN_VALIDRESP'));
         FusionFille.PutValue('PCN_EXPORTOK',TA.getValue('PCN_EXPORTOK'));
         FusionFille.PutValue('PCN_LIBCOMPL1',TA.getValue('PCN_LIBCOMPL1'));
         FusionFille.PutValue('PCN_LIBCOMPL2',TA.getValue('PCN_LIBCOMPL2'));
         FusionFille.PutValue('PCN_VALIDABSENCE',TA.getValue('PCN_VALIDABSENCE'));
         FusionFille.PutValue('PCN_OKFRACTION',TA.getValue('PCN_OKFRACTION'));
         FusionFille.PutValue('PSA_DATEANCIENNETE',TA.getValue('PSA_DATEANCIENNETE')); // PT3

// f PT2
         FusionFille.PutValue('PCN_NBJCARENCE',TA.getValue('PCN_NBJCARENCE')); // Nbre jours carence prévus
         FusionFille.PutValue('PCN_NBJCALEND',TA.getValue('PCN_NBJCALEND'));// Nbre jours calendaires absence
         FusionFille.PutValue('PCN_NBJIJSS',TA.getValue('PCN_NBJIJSS'));// Nbre jours IJSS prévus
         FusionFille.PutValue('PCN_DATEPAIEMENT', TA.getValue('PCN_DATEPAIEMENT'));// Date d'intégration de l'absence dans le bulletin
         FusionFille.PutValue('PCN_IJSSSOLDEE',TA.getValue('PCN_IJSSSOLDEE'));// IJSS Soldées
       end
// d PT1
       else
         // uniquement les absence non réglées
         FusionFille.free;
//f PT1
     end
     else
     if not DateRegltON then
     // pas de règlement pour cette absence
     begin
       FusionFille := TOB.create('REGLTIJSS', T_Reglt, -1);

       FusionFille.AddChampSup('LIBELLE',TRUE);
// d PT2
         FusionFille.AddChampSup('PCN_TYPEMVT',TRUE);
         FusionFille.AddChampSup('PCN_RESSOURCE',TRUE);
         FusionFille.AddChampSup('PCN_DATEDEBUT',TRUE);
         FusionFille.AddChampSup('PCN_DATEFIN',TRUE);
         FusionFille.AddChampSup('PCN_TYPECONGE',TRUE);
         FusionFille.AddChampSup('PCN_SENSABS',TRUE);
         FusionFille.AddChampSup('PCN_MVTORIGINE',TRUE);
         FusionFille.AddChampSup('PCN_TYPEIMPUTE',TRUE);
         FusionFille.AddChampSup('PCN_DATECREATION',TRUE);
         FusionFille.AddChampSup('PCN_LIBELLE',TRUE);
         FusionFille.AddChampSup('PCN_GENERECLOTURE',TRUE);
         FusionFille.AddChampSup('PCN_DATEMODIF',TRUE);
         FusionFille.AddChampSup('PCN_DATESOLDE',TRUE);
         FusionFille.AddChampSup('PCN_DATEVALIDITE',TRUE);
         FusionFille.AddChampSup('PCN_DATEDEBUTABS',TRUE);
         FusionFille.AddChampSup('PCN_DEBUTDJ',TRUE);
         FusionFille.AddChampSup('PCN_DATEFINABS',TRUE);
         FusionFille.AddChampSup('PCN_FINDJ',TRUE);
         FusionFille.AddChampSup('PCN_DATEPAIEMENT',TRUE);
         FusionFille.AddChampSup('PCN_JOURS',TRUE);
         FusionFille.AddChampSup('PCN_HEURES',TRUE);
         FusionFille.AddChampSup('PCN_BASE',TRUE);
         FusionFille.AddChampSup('PCN_NBREMOIS',TRUE);
         FusionFille.AddChampSup('PCN_CODERGRPT',TRUE);
         FusionFille.AddChampSup('PCN_MVTDUPLIQUE',TRUE);
         FusionFille.AddChampSup('PCN_ABSENCE',TRUE);
         FusionFille.AddChampSup('PCN_ABSENCEMANU',TRUE);
         FusionFille.AddChampSup('PCN_MODIFABSENCE',TRUE);
         FusionFille.AddChampSup('PCN_APAYES',TRUE);
         FusionFille.AddChampSup('PCN_VALOX',TRUE);
         FusionFille.AddChampSup('PCN_VALOMS',TRUE);
         FusionFille.AddChampSup('PCN_VALORETENUE',TRUE);
         FusionFille.AddChampSup('PCN_VALOMANUELLE',TRUE);
         FusionFille.AddChampSup('PCN_MODIFVALO',TRUE);
         FusionFille.AddChampSup('PCN_PERIODEPAIE',TRUE);
         FusionFille.AddChampSup('PCN_ETABLISSEMENT',TRUE);
         FusionFille.AddChampSup('PCN_TRAVAILN1',TRUE);
         FusionFille.AddChampSup('PCN_TRAVAILN2',TRUE);
         FusionFille.AddChampSup('PCN_TRAVAILN3',TRUE);
         FusionFille.AddChampSup('PCN_TRAVAILN4',TRUE);
         FusionFille.AddChampSup('PCN_CODESTAT',TRUE);
         FusionFille.AddChampSup('PCN_CONFIDENTIEL',TRUE);
         FusionFille.AddChampSup('PCN_SAISIEDEPORTEE',TRUE);
         FusionFille.AddChampSup('PCN_VALIDSALARIE',TRUE);
         FusionFille.AddChampSup('PCN_VALIDRESP',TRUE);
         FusionFille.AddChampSup('PCN_EXPORTOK',TRUE);
         FusionFille.AddChampSup('PCN_LIBCOMPL1',TRUE);
         FusionFille.AddChampSup('PCN_LIBCOMPL2',TRUE);
         FusionFille.AddChampSup('PCN_VALIDABSENCE',TRUE);
         FusionFille.AddChampSup('PCN_OKFRACTION',TRUE);
         FusionFille.AddChampSup('PSA_DATEANCIENNETE',TRUE);  // PT3

// f PT2
       FusionFille.AddChampSup('PCN_NBJCARENCE',TRUE);
       FusionFille.AddChampSup('PCN_NBJCALEND',TRUE);
       FusionFille.AddChampSup('PCN_NBJIJSS',TRUE);
       FusionFille.AddChampSup('PCN_IJSSSOLDEE',TRUE);
       FusionFille.AddChampSup('PCN_DATEPAIEMENT',TRUE);
       FusionFille.AddChampSup('LIBTYPEABS',TRUE);
       FusionFille.AddChampSup('MATRICULENOMPRENOM',TRUE);

       FusionFille.PutValue('PRI_SALARIE', TA.GetValue('PCN_SALARIE')); // matricule
       FusionFille.PutValue('MATRICULENOMPRENOM',Trim(FusionFille.GetValue('PRI_SALARIE'))+ // matricule + nom +prénom salarié
                                        '  '+
                                        Trim(TA.GetValue('PSA_LIBELLE'))+
                                        '  '+
                                        Trim(TA.GetValue('PSA_PRENOM'))); // nom salarié
       FusionFille.PutValue('LIBELLE', TA.getValue('PCN_LIBELLE')); // libellé absence
// d PT2
         FusionFille.PutValue('PCN_TYPEMVT',TRUE);
         FusionFille.PutValue('PCN_RESSOURCE',TRUE);
         FusionFille.PutValue('PCN_DATEDEBUT',TRUE);
         FusionFille.PutValue('PCN_DATEFIN', TA.getValue('PCN_DATEFIN'));
         FusionFille.PutValue('PCN_TYPECONGE', TA.getValue('PCN_TYPECONGE'));
         FusionFille.PutValue('PCN_SENSABS', TA.getValue('PCN_SENSABS'));
         FusionFille.PutValue('PCN_MVTORIGINE', TA.getValue('PCN_MVTORIGINE'));
         FusionFille.PutValue('PCN_TYPEIMPUTE', TA.getValue('PCN_TYPEIMPUTE'));
         FusionFille.PutValue('PCN_DATECREATION', TA.getValue('PCN_DATECREATION'));
         FusionFille.PutValue('PCN_LIBELLE', TA.getValue('PCN_LIBELLE'));
         FusionFille.PutValue('PCN_GENERECLOTURE', TA.getValue('PCN_GENERECLOTURE'));
         FusionFille.PutValue('PCN_DATEMODIF', TA.getValue('PCN_DATEMODIF'));
         FusionFille.PutValue('PCN_DATESOLDE', TA.getValue('PCN_DATESOLDE'));
         FusionFille.PutValue('PCN_DATEVALIDITE', TA.getValue('PCN_DATEVALIDITE'));
         FusionFille.PutValue('PCN_DATEDEBUTABS',TRUE);
          FusionFille.PutValue('PCN_DEBUTDJ', TA.getValue('PCN_DEBUTDJ'));
         FusionFille.PutValue('PCN_DATEFINABS', TA.getValue('PCN_DATEFINABS'));
         FusionFille.PutValue('PCN_FINDJ', TA.getValue('PCN_FINDJ'));
         FusionFille.PutValue('PCN_DATEPAIEMENT', TA.getValue('PCN_DATEPAIEMENT'));
         FusionFille.PutValue('PCN_JOURS', TA.getValue('PCN_JOURS'));
         FusionFille.PutValue('PCN_HEURES', TA.getValue('PCN_HEURES'));
         FusionFille.PutValue('PCN_BASE', TA.getValue('PCN_BASE'));
         FusionFille.PutValue('PCN_NBREMOIS', TA.getValue('PCN_NBREMOIS'));
         FusionFille.PutValue('PCN_CODERGRPT', TA.getValue('PCN_CODERGRPT'));
         FusionFille.PutValue('PCN_MVTDUPLIQUE', TA.getValue('PCN_MVTDUPLIQUE'));
         FusionFille.PutValue('PCN_ABSENCE', TA.getValue('PCN_ABSENCE'));
         FusionFille.PutValue('PCN_ABSENCEMANU', TA.getValue('PCN_ABSENCEMANU'));
         FusionFille.PutValue('PCN_MODIFABSENCE', TA.getValue('PCN_MODIFABSENCE'));
         FusionFille.PutValue('PCN_APAYES', TA.getValue('PCN_APAYES'));
         FusionFille.PutValue('PCN_VALOX', TA.getValue('PCN_VALOX'));
         FusionFille.PutValue('PCN_VALOMS', TA.getValue('PCN_VALOMS'));
         FusionFille.PutValue('PCN_VALORETENUE', TA.getValue('PCN_VALORETENUE'));
         FusionFille.PutValue('PCN_VALOMANUELLE', TA.getValue('PCN_VALOMANUELLE'));
         FusionFille.PutValue('PCN_MODIFVALO', TA.getValue('PCN_MODIFVALO'));
         FusionFille.PutValue('PCN_PERIODEPAIE', TA.getValue('PCN_PERIODEPAIE'));
         FusionFille.PutValue('PCN_ETABLISSEMENT', TA.getValue('PCN_ETABLISSEMENT'));
         FusionFille.PutValue('PCN_TRAVAILN1', TA.getValue('PCN_TRAVAILN1'));
         FusionFille.PutValue('PCN_TRAVAILN2', TA.getValue('PCN_TRAVAILN2'));
         FusionFille.PutValue('PCN_TRAVAILN3', TA.getValue('PCN_TRAVAILN3'));
         FusionFille.PutValue('PCN_TRAVAILN4', TA.getValue('PCN_TRAVAILN4'));
         FusionFille.PutValue('PCN_CODESTAT', TA.getValue('PCN_CODESTAT'));
         FusionFille.PutValue('PCN_CONFIDENTIEL', TA.getValue('PCN_CONFIDENTIEL'));
         FusionFille.PutValue('PCN_SAISIEDEPORTEE', TA.getValue('PCN_SAISIEDEPORTEE'));
         FusionFille.PutValue('PCN_VALIDSALARIE', TA.getValue('PCN_VALIDSALARIE'));
         FusionFille.PutValue('PCN_VALIDRESP', TA.getValue('PCN_VALIDRESP'));
         FusionFille.PutValue('PCN_EXPORTOK', TA.getValue('PCN_EXPORTOK'));
         FusionFille.PutValue('PCN_LIBCOMPL1', TA.getValue('PCN_LIBCOMPL1'));
         FusionFille.PutValue('PCN_LIBCOMPL2', TA.getValue('PCN_LIBCOMPL2'));
         FusionFille.PutValue('PCN_VALIDABSENCE', TA.getValue('PCN_VALIDABSENCE'));
         FusionFille.PutValue('PCN_OKFRACTION', TA.getValue('PCN_OKFRACTION'));
         FusionFille.PutValue('PSA_DATEANCIENNETE', TA.getValue('PSA_DATEANCIENNETE'));  // PT3

// f PT2
       FusionFille.PutValue('PCN_NBJCARENCE',TA.getValue('PCN_NBJCARENCE')); // Nbre jours carence prévus
       FusionFille.PutValue('PCN_NBJCALEND',TA.getValue('PCN_NBJCALEND'));// Nbre jours calendaires absence
       FusionFille.PutValue('PCN_NBJIJSS',TA.getValue('PCN_NBJIJSS'));// Nbre jours IJSS prévus
       FusionFille.PutValue('PCN_DATEPAIEMENT', TA.getValue('PCN_DATEPAIEMENT'));// Date d'intégration de l'absence dans le bulletin
       FusionFille.PutValue('PCN_IJSSSOLDEE',TA.getValue('PCN_IJSSSOLDEE'));// IJSS Soldées
       FusionFille.PutValue('PRI_DATEDEBUTABS', TA.GetValue('PCN_DATEDEBUTABS'));// Date de début d'absence
       FusionFille.PutValue('PRI_DATEFINABS', TA.GetValue('PCN_DATEFINABS'));// Date de fin d'absence
       FusionFille.PutValue('PRI_DATEINTEGR',IDate1900);// Date d'intégration des IJSS dans le bulletin // PT9
       FusionFille.PutValue('PRI_DATEREGLT',IDate1900);// Date de versement des IJSS // PT9
       FusionFille.PutValue('PRI_MTREGLE',0);// Mt IJSS nettes
       FusionFille.PutValue('PRI_NBJCARENCE',0);// Nbre de jours de carence IJSS
       FusionFille.PutValue('PRI_MTIJSSBRUTES',0); // mt IJSS brutes
       FusionFille.PutValue('PRI_NBJIJSS', 0);// Nbre de jours d'IJSS
       Q := OpenSql('SELECT PMA_TYPEABS FROM MOTIFABSENCE '+
                    'WHERE ##PMA_PREDEFINI## PMA_MOTIFABSENCE = "'+  { PT7 }
                    TA.getvalue('PCN_TYPECONGE') + '"',TRUE);
       if not Q.EOF then
         FusionFille.PutValue('PRI_TYPEABS',Q.FindField('PMA_TYPEABS').AsString);
       ferme (Q);
       FusionFille.PutValue('LIBTYPEABS',RechDom('PGTYPEABS',FusionFille.GetValue('PRI_TYPEABS'),FALSE));

     end;
   end;

     TFStat(Ecran).LaTOB := T_Reglt;
end;
// d PT1
procedure TOF_STAREGLTIJ.ExitEdit(Sender: TObject);
var
  edit                                          : thedit;

begin
  edit:=THEdit(Sender);
  if edit <> nil then
      if (edit.text <> '  /  /    ')  then
      if (not IsValidDate(edit.text)) then
      SetFocusControl('DATEDEBREGLT');

end;  { fin ExitEdit}
procedure TOF_STAREGLTIJ.ChangeChk(Sender: TObject);
var
  Chk, ChkRgl                                          : TCheckBox;
begin
  Chk:=TCheckBox(Sender);
  if Chk <> nil then
  begin
    if (Chk.name = 'CHKREGLE')  then
    begin
      if (Chk.checked = true) then
      begin
        ChkRgl := TCheckBox(getcontrol('CHKREGLE1'));
        ChkRgl.checked := false;
      end;
    end;
    if (Chk.name = 'CHKREGLE1')  then
    begin
      if (Chk.checked = true) then
      begin
        ChkRgl := TCheckBox(getcontrol('CHKREGLE'));
        ChkRgl.checked := false;
      end;
    end;
  end;
end;  { fin ChangeChk}
// f PT1
// d PT4
procedure TOF_STAREGLTIJ.ExitEditS(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);

if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and
       (length(Edit.text)<11) and
       (isnumeric(edit.text)) then
      edit.text:=AffectDefautCode(edit,10);
end;  { fin ExitEditS}
// f PT4
Initialization
  registerclasses ( [ TOF_STAREGLTIJ ] ) ;
end.
