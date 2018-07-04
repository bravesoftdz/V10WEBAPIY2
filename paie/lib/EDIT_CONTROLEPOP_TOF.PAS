{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 07/09/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : EDIT_CONTROLEPOP ()
                 Contrôle des populations
Mots clefs ... : TOF;EDIT_CONTROLEPOP
*****************************************************************}
Unit EDIT_CONTROLEPOP_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul,
     qrs1,
{$else}
     eMul,
     eqrs1,
    {$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     HQry,
     Utob,
     pgpopuloutils,
     pgoutils2,
     UTobDebug,
     Entpaie,
     UTOF ; 

Type
  TOF_EDIT_CONTROLEPOP = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
    tobedit : tob;
    procedure change(sender:TObject);
  end ;

Implementation

procedure TOF_EDIT_CONTROLEPOP.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_EDIT_CONTROLEPOP.OnDelete ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 07/09/2006
Modifié le ... :   /  /    
Description .. : Edition
Mots clefs ... : 
*****************************************************************}
procedure TOF_EDIT_CONTROLEPOP.OnUpdate ;
var
  checkdoublon, checkexclu : Tcheckbox;
  ch_choixetat : THedit;
  QQ, Q1 : TQuery;
  datectl : Tdatetime;
  T1,TD, tob_popsalfille, tob_popsal, tob_sal, tob_salfille, Tob_pop, tob_popfille: Tob;
  salarie, typepop, population, typeetat, champbl, libpop, codepop , sql1, nodossier: string;
  editsal, trtok : boolean;
  i, k : integer;

begin
  Inherited ;
  trtok := true;
  checkdoublon := Tcheckbox(Getcontrol('DOUBLON'));
  checkexclu := Tcheckbox(Getcontrol('EXCLUS'));

  typepop := THValcomboBox(getcontrol('TYPEPOPUL')).value;
  if typepop = '' then
  begin
  pgibox('Vous devez renseigner le type de la population.', TFQRS1(Ecran).Caption);
  setfocuscontrol('TYPEPOPUL');
  trtok := false;
  end
  else
  begin
  if (checkdoublon.checked = false) and (checkexclu.checked = false) then
  begin
  pgibox('Choisir le type de l''état  : doublons ou exclus.', TFQRS1(Ecran).Caption);
  setfocuscontrol('DOUBLON');
  trtok := false;
  end;
  end;

  champbl := '';

   // chargement des données de l'édition
  ch_choixetat := ThEdit(Getcontrol('CHOIXETAT'));
  if Trim(ch_choixetat.text) <> '' then typeetat := trim(ch_choixetat.text);

  datectl :=  StrToDate(getcontroltext('DATEPOPUL'));

  if assigned(TobEdit) then FreeAndNil(TobEdit);
  Tobedit := Tob.Create('Editions',nil,-1);


  if (trtok = true) then
  begin
    nodossier := pgrendNodossier();
  // chargement des données population dans la tob  TOB_POPSAL  (population, libellé, salarié affecté)
  // recherche des populations de type DOSSIER
  sql1 := 'select PPC_POPULATION,PPC_LIBELLE,PPC_CODEPOP, ppo_typepop  from codepopulation,ordrepopulation '+
  ' where PPO_PREDEFINI= "DOS" and PPO_NODOSSIER = "'+nodossier+'" and ppc_codepop=ppo_codepop and'+
  ' ppo_typepop like "%'+ typepop +'%"';
    if not existesql(sql1) then
    begin
      // recherche des populations de type Standard
      sql1 := 'select PPC_POPULATION,PPC_LIBELLE,PPC_CODEPOP, ppo_typepop  from codepopulation,ordrepopulation '+
      ' where PPC_PREDEFINI= "STD" and ppc_codepop=ppo_codepop and ppo_typepop like "%'+ typepop +'%" ';
      if not existesql(sql1) then
      // recherche des populations de type CEGID
      sql1 := 'select PPC_POPULATION,PPC_LIBELLE,PPC_CODEPOP, ppo_typepop  from codepopulation,ordrepopulation '+
      ' where PPC_PREDEFINI= "CEG" and ppc_codepop=ppo_codepop and ppo_typepop like "%'+ typepop +'%"';
    end;

  Q1:= Opensql(sql1, true);

    if tob_popsal <> nil then freeandnil(tob_popsal);
    Tob_popsal := Tob.Create('population_salarie',nil,-1);
    While Not Q1.Eof Do
    begin    // pour chaque population: lecture des salariés affectés qui lui sont affectés (tob_sal)

    population := Q1.findfield('PPC_POPULATION').asstring;
    codepop := Q1.findfield('PPC_CODEPOP').asstring;
    if tob_sal <> nil then freeandnil(tob_sal);
    tob_sal := salariepop(codepop,population,datectl);

          Tob_salfille := TOB_sal.FindFirst([''], [''], TRUE);

          while Tob_salfille <> nil do
          begin
          tob_popsalfille := Tob.Create('Population_salarie',Tob_popsal,-1);
          Tob_popsalfille.AddChampSupValeur('POPULATION',population);
          Tob_popsalfille.AddChampSupValeur('LIBELLE',Q1.findfield('PPC_LIBELLE').asstring);
          tob_popsalfille.Addchampsupvaleur('SALARIE',Tob_salfille.getvalue('SALARIE'));
          Tob_salfille := TOB_sal.FindNext([''], [''], TRUE);
          end;
          if tob_sal <> nil then freeandnil(tob_sal);
    Q1.next;
    end;
    ferme(Q1);

  // recherche des salariés présents à la date saisie
  QQ:= opensql('select PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM, PSA_ETABLISSEMENT, PSA_DATEENTREE from SALARIES where'+
  ' (PSA_DATESORTIE >="'+usdatetime(datectl)+'" or'+
  ' PSA_DATESORTIE <="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL) and (PSA_DATEENTREE <= "'+
  usdatetime(datectl)+'")', true);

  While Not QQ.Eof Do
  begin   // d1
   editsal := false;
   salarie := QQ.findfield('PSA_SALARIE').asstring;

   Tob_pop := Tob.Create('Populations du salarie',nil,-1);
   TD := Tob_popsal.findfirst(['SALARIE'], [salarie], true);
    // pour chaque salarié ==> création de la tob tob_pop qui contient les populations auxquelles il est affecté
     while TD <> nil do
     begin
     Tob_popfille := Tob.Create('populations du salarie',Tob_pop,-1);
     population := TD.getvalue('POPULATION');
     Tob_popfille.AddChampSupValeur('POPULATION',population);
     libpop := TD.getvalue('LIBELLE');
     Tob_popfille.AddChampSupValeur('LIBELLE',libpop);
     TD:= Tob_popsal.findNext(['SALARIE'], [salarie], true);
     end;

    // salarié sans population
    if (Tob_pop.detail.count = 0) and (typeetat = 'EXCLUS')  then editsal := true;

    // salarié appartenant à une population
    if (Tob_pop.detail.count > 1) and (typeetat = 'DOUBLON') then editsal := true;

    // alimente la tob d'édition
    if (editsal = true) then
    begin  // d2
    T1 := Tob.Create('Edit',TobEdit,-1);
    T1.AddChampSupValeur('PSA_SALARIE',salarie);
    T1.AddChampSupValeur('PSA_LIBELLE',QQ.findfield('PSA_LIBELLE').asstring);
    T1.AddChampSupValeur('PSA_PRENOM', QQ.findfield('PSA_PRENOM').asstring);
    T1.AddChampSupValeur('PSA_ETABLISSEMENT',QQ.findfield('PSA_ETABLISSEMENT').asstring);
    T1.AddChampSupValeur('PSA_DATEENTREE',QQ.findfield('PSA_DATEENTREE').asstring);

       if  typeetat = 'DOUBLON'  then
       begin  // d3

       if TOB_pop <> nil then
       begin  // d4
         for k := 1 to 5 do
         begin
         T1.AddChampSupValeur('POPUL'+ IntToStr(k),champbl);
         T1.AddChampSupValeur('LIBPOPUL'+IntToStr(k),champbl);
         end;

          // alimente les 5 premières populations auxquelles appartient le salarié
          Tob_popfille := TOB_pop.FindFirst([''], [''], TRUE);
          i := 0 ;
          while Tob_popfille <> nil do
          begin  // d5
          i := i + 1;
          if i > 5 then break;
          population := Tob_popfille.getvalue('POPULATION');
          libpop := '';
          libpop := tob_popfille.getvalue('LIBELLE');
          T1.putvalue('POPUL'+ IntToStr(i),population);
          T1.putvalue('LIBPOPUL'+IntToStr(i), libpop);
          Tob_popfille := TOB_pop.FindNext([''], [''], TRUE);
          end; // e5


       TFQRS1(Ecran).CodeEtat:= 'SAP';
      end; // e4
      end // e3  
      else
      TFQRS1(Ecran).CodeEtat:= 'SSP';

    end; // e2
    if assigned(Tob_pop) then FreeAndNil(Tob_pop);
    QQ.next;
  end; // e1
  ferme(QQ);
end;

  if assigned(Tob_popsal) then FreeAndNil(Tob_popsal);
 // tobdebug(tobedit);
  TobEdit.detail.sort('PSA_ETABLISSEMENT;PSA_SALARIE');
  TFQRS1(Ecran).LaTob:= TobEdit;
 end ;


procedure TOF_EDIT_CONTROLEPOP.OnLoad ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 08/09/2006
Modifié le ... :   /  /    
Description .. : On argument
Mots clefs ... : 
*****************************************************************}
procedure TOF_EDIT_CONTROLEPOP.OnArgument (S : String ) ;
var
check : tcheckbox;
Stplus : string;

begin
  Inherited ;

  stplus := '';
  if not VH_PAIE.PGSeriaFormation then stplus := ' AND CO_CODE <> "FOR"';
  if not VH_PAIE.PGSeriaPresence then  stplus := stplus + ' AND CO_CODE <> "PRE"';
  if stplus <> '' then SetControlProperty('TYPEPOPUL', 'Plus', stplus);


  check := tcheckbox(getcontrol('DOUBLON'));
  if check <> nil then check.onclick := change;
  check := tcheckbox(getcontrol('EXCLUS'));
  if check <> nil then check.onclick := change;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 11/09/2006
Modifié le ... :   /  /    
Description .. : On close
Mots clefs ... : 
*****************************************************************}
procedure TOF_EDIT_CONTROLEPOP.OnClose ;
Var
  DateDefaut : THEdit;
begin
  Inherited ;
  if assigned(TobEdit) then FreeAndNil(TobEdit);

  DateDefaut := THEdit(GetControl('DATEPOPUL'));
  If DateDefaut <> nil Then
     If (not(IsValidDate(DateDefaut.text))) Then
        SetControltext('DATEPOPUL',DateToStr(Date));


end ;

procedure TOF_EDIT_CONTROLEPOP.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_EDIT_CONTROLEPOP.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 07/09/2006
Modifié le ... :   /  /    
Description .. : Click sur doublon ou sur exclu
Mots clefs ... : 
*****************************************************************}
procedure TOF_EDIT_CONTROLEPOP.change(sender : Tobject);
var
ch_choixetat : Thedit;
checkdoublon, checkexclu : Tcheckbox;

begin
ch_choixetat := THedit(getcontrol('CHOIXETAT'));
checkdoublon := Tcheckbox(Getcontrol('DOUBLON'));
checkexclu := Tcheckbox(Getcontrol('EXCLUS'));


if ch_choixetat <> nil then
begin
 checkexclu.enabled := true;
 checkdoublon.enabled := true;
 ch_choixetat.text := '';

 if (checkdoublon <> nil) and (checkdoublon.checked = true) then begin
 ch_choixetat.text := 'DOUBLON';
 checkexclu.enabled := false;
 end;

 if (checkexclu <> nil) and (checkexclu.checked = true) then begin
 ch_choixetat.text := 'EXCLUS';
 checkdoublon.enabled := false;
 end;

 end;
end;

Initialization
  registerclasses ( [ TOF_EDIT_CONTROLEPOP ] ) ; 
end.
