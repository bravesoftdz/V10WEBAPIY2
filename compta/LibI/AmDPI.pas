{***********UNITE*************************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 03/05/2006
Modifié le ... :   /  /
Description .. : Gestion de la class TAmDPI afférente au suivi des
Suite ........ : Déductions pour investissement
Mots clefs ... :
*****************************************************************}
unit AmDPI;


interface

uses classes, HEnt1, Hctrls, UTob,
     SysUtils
     // BTY 07/06 Pour compilation SERIE1
     {$IFDEF SERIE1}
     ,S1Util, utModules
     {$ELSE}
     {$IFDEF MODENT1}
     ,CPTypeCons
     {$ENDIF MODENT1}
     ,Ent1
     {$ENDIF}
     {$IFNDEF EAGLCLIENT}
     {$IFDEF ODBCDAC}
      , odbcconnection, odbctable, odbcquery, odbcdac
      {$ELSE}
       {$IFNDEF DBXPRESS}
       ,dbtables
       {$ELSE}
       ,uDbxDataSet
       {$ENDIF}
      {$ENDIF}
     {$ENDIF}
     ;

type
  TAmDPI = class
  private
  public
       Exercice: array[0..5] of TExoDate;
       Solde: array[0..5] of double;
       Montant : double;
       SoldeDispoDate : double;
       Anterieur : double;
       St : string;
       constructor Create;
       destructor Destroy; override;
       procedure  ChargeDateDPI ;
       procedure  RAZSoldeDPI ;
       procedure  ChargeDateEtSoldeDPI ;
       procedure  ChargeSoldeDisponibleDate (FinExoDPI : TDateTime);
  end;


implementation

uses  ImEnt
      , Outils
      ;


constructor TAmDPI.Create;
var i: integer;
begin
  inherited Create;
  for i:= 0 to 5 do
  begin
    Exercice[i].Deb := iDate1900;
    Exercice[i].Fin := iDate1900;
  end;
end;

destructor TAmDPI.Destroy;
begin
  inherited Destroy;
end;

procedure  TAmDPI.ChargeDateDPI ;

var Premier: boolean ;
    iEncours,i,j: integer ;
    TOBExo : Tob;
    Q: TQuery ;

begin
  // Récup exercices dans une TOB
  TOBExo := TOB.Create ('',nil,-1);
  Q := OpenSQL('SELECT EX_DATEDEBUT,EX_DATEFIN,EX_EXERCICE,EX_ETATCPTA,EX_LIBELLE FROM EXERCICE ORDER BY EX_DATEDEBUT',TRUE) ;
  TOBExo.LoadDetailDB('', '', '', Q, false);
  Ferme(Q);

  // Récup rang de l'exercice en cours
  iEncours := 0;
  Premier:=true;
  for i:=0 to TOBExo.Detail.Count - 1 do
  begin
    if (TOBExo.Detail[i].GetValue('EX_ETATCPTA')= 'OUV') or
       (TOBExo.Detail[i].GetValue('EX_ETATCPTA')= 'CPR') then
      if Premier then
         begin
         iEncours := i;
         Break;
         end;
  end;

  // Alimenter tableau Exercice en reculant les dates à partir de l'exo en cours
  i := iEncours;
  j := 0;
  if TOBExo.Detail.Count <> 0 then
  begin
    while (TOBExo.Detail[i].GetValue ('EX_EXERCICE') <> '') do
    begin
      Exercice[j].Code := TOBExo.Detail[i].GetValue ('EX_EXERCICE');
      Exercice[j].Deb := TOBExo.Detail[i].GetValue ('EX_DATEDEBUT');
      Exercice[j].Fin := TOBExo.Detail[i].GetValue ('EX_DATEFIN');
      Exercice[j].EtatCpta := TOBExo.Detail[i].GetValue ('EX_LIBELLE');
      i := i-1;
      Inc(j);
      // Arrêter si TOBExo.Detail[i] tout déroulé ou si 6 exercices alimentés
      if (i < 0) or (j > 5) then Break;
    end;
  end;

  // Vérif présence des 6 exercices et calcul exos si non renseignés
  for j:= 0 to 5 do
  begin
    if (Exercice[j].Code = '') and (j > 0) then
       if Exercice[j-1].Deb = iDate1900 then
          // Date début exo d'avant non renseigné, en principe impossible
          Break
       else
          begin
          Exercice[j].Code := '';
          Exercice[j].Fin := Exercice[j-1].Deb-1;
          if Exercice[j].Fin = FinDeMois (Exercice[j].Fin) then
            // assurer le 29 fév sur une année bissextile
            Exercice[j].Deb := FinDeMois(PlusMois (Exercice[j].Fin, -12)) + 1
          else
            Exercice[j].Deb := PlusMois (Exercice[j].Fin, -12)+ 1;
          end;
  end;

  TOBExo.Free;
end;


procedure  TAmDPI.RAZSoldeDPI ;
var i : integer;
begin
  for i:=0 to 5 do
     Solde [i] := 0;
end;


procedure  TAmDPI.ChargeDateEtSoldeDPI ;
var Q: TQuery;
    TMvtD, T : Tob;
    i, j: integer;
    mnt, util : double;
begin

   TMvtD := TOB.Create ('',nil,-1);

   // Mise en TOB de la table IMMOMVTD par date décroissante
   Q := OpenSQL('SELECT IZ_DATE,IZ_SEQUENCE,IZ_MONTANT,IZ_NATURED,IZ_IMMO,IZ_LIBELLE FROM IMMOMVTD ORDER BY IZ_DATE DESC', True);
   TMvtD.LoadDetailDB('', '', '', Q, false);
   Ferme(Q);

   RAZSoldeDPI;
   ChargeDateDPI;
   if (TMvtD.Detail.Count = 0) then exit;

   // Commencer au 1er exercice du tableau, c'est fAmDPI.Exercice[0]
   // car on peut avoir saisi de la DPI sur un exo dont on a ensuite annulé la clôture
   i := 0;
   while (TMvtD.Detail[i].GetValue('IZ_DATE') > Exercice[0].Fin) do
     begin
     Inc (i);
     if i = TMvtD.Detail.Count then  Break;
     end;

   // 1e Tob fille
   if (i < TMvtD.Detail.Count) then T := TMvtD.Detail[i]
   else exit;

   // Boucler sur chaque exercice j
   for j:=0 to 5 do
   begin
     mnt := Arrondi (0.00, V_PGI.OkDecV);
     util := Arrondi (0.00, V_PGI.OkDecV);

     // Tous les enregs du même exo j
     while (T.GetValue('IZ_DATE') = Exercice[j].Fin) do
     begin
        if (T.GetValue('IZ_NATURED')='DPI')  then
           // montants de la DPI de l'exo j
           mnt := Arrondi (Mnt + T.GetValue('IZ_MONTANT'),V_PGI.OkDecV)
        else
           // utilisations
           util := Arrondi (Util + T.GetValue('IZ_MONTANT'),V_PGI.OkDecV);

        // Enreg suivant
        Inc(i);
        if (i < TMvtD.Detail.Count) then T := TMvtD.Detail[i]
        else  T:= nil;
        if (T=nil) then Break;
     end; // while

     // Solde de l'exo j
     if j=0 then
        Solde [j] := Arrondi (mnt, V_PGI.OkDecV)
     else
        Solde [j] := Arrondi ((mnt-util), V_PGI.OkDecV);

     // Sortir de la boucle si on a tout déroulé TMvtD
     if (T=nil) then Break;
   end; // for

   TMvtD.Free;
end;

procedure TAmDPI.ChargeSoldeDisponibleDate (FinExoDPI : TDateTime);
var i, j:integer;
begin
  ChargeDateEtSoldeDPI;
  SoldeDispoDate := 0;
  i :=0;

  // Indice de l'exo passé en entrée
  while Exercice[i].Fin > FinExoDPI do
    begin
    Inc (i);
    if i = 6 then Break;
    end;
  if (i = 6) then exit;

  // Commencer à l'exo précédant celui en entrée
  for j := i+1 to 5 do
    begin
      SoldeDispoDate := Arrondi (SoldeDispoDate + Solde [j], V_PGI.OkDecV);
    end;
end;


end.
