{***********UNITE*************************************************
Auteur  ...... :  NA
Créé le ...... : 28/03/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : SIMULREGLESPOP ()
                 Calcul de la simulation IDR
Mots clefs ... : TOF;CALCULSIMUL
*****************************************************************
Pt1      : 17/04/2007    NA    Affichage des populations DOS ou STD ou CEG}

Unit CALCULSIMUL_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     HDB,
    {$else}
     eMul,
     {$ENDIF}
     uTob,
     ed_tools,
     HStatus,
     sysutils,
     ComCtrls,
     HCtrls,
     HMsgBox,
     HTB97,
     HQRY,
     UTobDebug,
     Pgoutils2,
     pgcalculsimulation,
     UTOF;


Type
  TOF_CALCULSIMUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

    Private
    {$IFNDEF EAGLCLIENT}
    Liste: THDBGrid;
{$ELSE}
    Liste: THGrid;
{$ENDIF}
    Paramsimul : Tparamsimul;
    mode, codesimd : string;
    procedure simulation(Sender : Tobject);
  end;

Implementation

Uses
  PGPOPULOUTILS;

procedure TOF_CALCULSIMUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CALCULSIMUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CALCULSIMUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CALCULSIMUL.OnLoad ;
var
listewhere, popul, selectpop, nodossier, pred : string;  //  , sql1
QQ : Tquery;

begin
  Inherited ;
  if mode = 'C' then
  begin
    // Pt1 recherche du prédéfini des populations IDR à afficher
    pred := GetPredefiniPopulation('IDR');
    nodossier := PgrendNodossier();
    // recherche des codes pop prédéfini DOSSIER
{    pred := 'DOS';
    sql1 := 'select PPC_POPULATION,PPO_CODEPOP from ORDREPOPULATION ,CODEPOPULATION where PPC_CODEPOP = PPO_CODEPOP'+
    ' and PPO_PREDEFINI= "DOS" and PPO_NODOSSIER = "'+nodossier+'" and PPO_TYPEPOP LIKE "%IDR%"';
    if not existesql(sql1) then
     begin
        // recherche des codes pop prédéfini STANDARD
        sql1 := 'select PPC_POPULATION,PPO_CODEPOP from ORDREPOPULATION,CODEPOPULATION where PPC_CODEPOP = PPO_CODEPOP'+
        ' and PPO_PREDEFINI= "STD" and PPO_TYPEPOP LIKE "%IDR%" ';
        pred := 'STD';
        if not existesql(sql1) then
        begin
          // recherche des codes pop prédéfini CEGID
          sql1 := 'select PPR_POPULATION, PPO_CODEPOP from ORDREPOPULATION, CODEPOPULATION where PPC_CODEPOP = PPO_CODEPOP'+
          ' and PPO_PREDEFINI= "CEG" and PPO_TYPEPOP LIKE "%IDR%" ';
          pred := 'CEG';
        end;
     end;  }
     // fin pt1
    SetControlText('XX_WHERE', '##PPR_PREDEFINI## PPC_PREDEFINI = "'+pred+'" and ((PPR_PREDEFINI= "CEG" AND PPR_POPULATION NOT IN'+
    ' (SELECT POP.PPR_POPULATION FROM REGLESPOP POP WHERE POP.PPR_PREDEFINI<>"CEG")) OR (PPR_PREDEFINI="STD" '+
    'AND PPR_POPULATION NOT IN (SELECT POP1.PPR_POPULATION FROM REGLESPOP POP1 WHERE POP1.PPR_PREDEFINI="DOS"))'+
    'OR (PPR_PREDEFINI="DOS"))');
  end
  else
  begin
    QQ:= opensql('Select PSL_SELECTPOP from SIMULATION where PSL_CODESIMUL = "'+codesimd+'"', true);
    if not (QQ.EOF) then
    begin
      listewhere := '';
      selectpop := QQ.findfield('PSL_SELECTPOP').asstring;
      popul := READTOKENST(selectpop);
      while popul <> '' do                                // lecture de chaque population
      begin
        listewhere :=  listewhere + '"'+popul+'" ,' ;
        popul := READTOKENST(selectpop);
      end;
    end;
    ferme(QQ);

    if listewhere <> '' then                  // si au moins une population
    listewhere := '(' + copy(listewhere, 1, length(listewhere) - 1) + ')';

    SetControlText('XX_WHERE','PPR_POPULATION in  '+listewhere+' and ((PPR_PREDEFINI= "CEG" AND PPR_POPULATION NOT IN'+
    ' (SELECT POP.PPR_POPULATION FROM REGLESPOP POP WHERE POP.PPR_PREDEFINI<>"CEG")) OR (PPR_PREDEFINI="STD" '+
    'AND PPR_POPULATION NOT IN (SELECT POP1.PPR_POPULATION FROM REGLESPOP POP1 WHERE POP1.PPR_PREDEFINI="DOS"))'+
    'OR (PPR_PREDEFINI="DOS")) ');
  end;
end ;

procedure TOF_CALCULSIMUL.OnArgument (S : String ) ;
var
Btn : TToolBarButton97;
codesimul, dtsimul , libellesimul, datecalcul, typedepart, methcalc, codecalc, txactu : string;


begin
  Inherited ;
  TFMul(Ecran).Retour := ' ';
  dtsimul := readtokenst(S);
  codesimul := readtokenst(S);
  libellesimul := readtokenst(S);
  datecalcul := readtokenst(S);
  typedepart := readtokenst(S);
  methcalc := readtokenst(S);
  codecalc := readtokenst(S);
  txactu := readtokenst(S);
  mode := readtokenst(S);
  codesimd := readtokenst(S);

  paramsimul.tdatesimul := dtsimul;
  paramsimul.tcodesimul := codesimul;
  paramsimul.Tlibellesimul:= libellesimul;
  paramsimul.Tdatecalcul := datecalcul;
  paramsimul.ttypedepart := typedepart;
  paramsimul.tmethcalcul := methcalc;
  paramsimul.Tcodecalcul:= codecalc;
  paramsimul.Ttauxactual := strtofloat(txactu);

  Ecran.caption := Ecran.caption + ' ' + codesimul + ' ' + libellesimul;

  Btn := TToolBarButton97(GetControl('B_CALCULSIMUL'));
  if btn <> nil then Btn.Onclick := simulation;

  if mode <> 'C' then begin
  setcontrolenabled('BselectAll', false);
  setcontrolvisible('BselectAll', false);
  end;

end ;

procedure TOF_CALCULSIMUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CALCULSIMUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CALCULSIMUL.OnCancel () ;
begin
  Inherited ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 10/07/2006
Modifié le ... :   /  /    
Description .. : SIMULATION
Mots clefs ... : 
*****************************************************************}
procedure TOF_CALCULSIMUL.Simulation (sender : Tobject);
var
pop , pred: string; // , sql1
i: integer;
tob_pop, T1 : Tob;
Q_Mul : THQUERY;

begin // D1

{$IFNDEF EAGLCLIENT}
Liste:= THDBGrid(GetControl('FListe'));
{$ELSE}
Liste:= THGrid(GetControl('FListe'));
{$ENDIF}

if Liste <> nil then
   begin  // D2
   if mode <> 'C' then Liste.Allselected := true;
   
   if (Liste.NbSelected = 0) and (not Liste.AllSelected) then
      begin
      PGIBOX('Aucune population sélectionnée', Ecran.Caption);
      exit;
      end;

  { Gestion de la sélection de population }

  if PgiAsk('Calcul de la simulation pour les populations concernées. Voulez-vous poursuivre ?', Ecran.caption) = mrYes then
    begin  // D3

    Q_Mul:= THQuery(Ecran.FindComponent('Q'));


      // traitement simulation

      Tob_pop := Tob.Create('Les populations',Nil,-1);
      i:=0;
      InitMoveProgressForm (NIL,'Traitement simulation en cours', 'Veuillez patienter SVP ...',i,FALSE,TRUE);
      InitMove(TFmul(Ecran).Q.RecordCount,'');

      if (Liste.AllSelected = TRUE) then
      begin   // D4
      // Si tout est sélectionné
      TFmul(Ecran).Q.First;
      while not TFmul(Ecran).Q.EOF do
            begin
             T1 := Tob.Create('Population',Tob_pop,-1);
             T1.AddChampSup('POPULATION',False);
             pop := Q_mul.findfield('PPR_POPULATION').asstring;
             T1.PutValue('POPULATION',pop);
             T1.AddChampSup('PREDEFINI',False);
             pred := Q_mul.findfield('PPR_PREDEFINI').asstring;
             T1.PutValue('PREDEFINI',pred);
             TFmul(Ecran).Q.Next;
            end;
      Liste.AllSelected := False;
      TFMul(Ecran).bSelectAll.Down := Liste.AllSelected;
      end  // E4
   else
     // lecture de chaque règle sélectionnée
      begin  // D5

      for i := 0 to Liste.NbSelected - 1 do
          begin
          {$IFDEF EAGLCLIENT}
          TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
          {$ENDIF}
          Liste.GotoLeBOOKMARK(i);
          T1 := Tob.Create('Population',Tob_pop,-1);
          T1.AddChampSup('POPULATION',False);
          pop := Q_mul.findfield('PPR_POPULATION').asstring;
          T1.PutValue('POPULATION',pop);
          T1.AddChampSup('PREDEFINI',False);
          pred:= Q_mul.findfield('PPR_PREDEFINI').asstring;
          T1.PutValue('PREDEFINI',pred);
          end;


      Liste.ClearSelected;
     end; // E5
     

     //tobdebug(Tob_pop);
    // Calcul de la simulatin avec en paramètre : TOB_POP : populations sélectionnées pour le calcul
    // et Paramsimul  : paramètres de calcul de la simulation
    Calculsimulation(Tob_pop, paramsimul);
    TFMul(Ecran).Retour := 'OK';
    

    FiniMove;
    FiniMoveProgressForm;
    Pgiinfo('Traitement terminé','Calcul de la simulation');
    TFMul(Ecran).BAnnulerClick(TFMul(Ecran).BAnnuler);

     end; // E3
    end; // E2
end; // E1

Initialization
  registerclasses ( [ TOF_CALCULSIMUL ] ) ;
end.
