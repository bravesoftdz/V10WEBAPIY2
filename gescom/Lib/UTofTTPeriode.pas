unit UTofTTPeriode;

interface

uses  SysUtils, Classes, Controls, Forms,
      {$IFNDEF EAGLCLIENT}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} db,
      {$ENDIF}
      ParamSoc, ed_tools, HMsgBox, HCtrls, HStatus, HEnt1, Ent1, UTOB, UTOF,
      InvUtil,UTofConsultStock, EntGC;


Type
     TOF_TTFinMois = Class (TOF)
        procedure OnArgument (stArgument : String ) ; override ;
        procedure OnUpdate ; override ;
        procedure OnClose ; override ;

        function  DatesOk : boolean;
        procedure InitDates;

        function TransactionFinMoisOK : boolean; //Recherche de la période à clôturer et appel du traitement
     end ;

Type
     TOF_TTFinExercice = Class (TOF)
        procedure OnArgument (stArgument : String ) ; override ;
        procedure OnUpdate ; override ;
        procedure OnClose ; override ;

        procedure InitDates;
        function  ControlDate(DateFinExo : TDateTime) : boolean;

        function TransactionFinExoOK : boolean; //Recherche de la période à clôturer et appel du traitement
     end ;

  //initialisation des variables
  procedure InitVar ;

  //Entrée du traitement suivant une date donnée
  procedure Entree_Traitement(DatClot : TDateTime ; FinExo : boolean);

  //MAJ de la TOB Dispo
  procedure MAJQuantite (TOBTT : TOB ; DatClot : TDateTime ; FinExo : boolean);
  procedure ClotureMois(TOBTT : TOB ; DatClot : TDateTime);

  //Fonctions de recherches de dates
  function  RechDerniereCloture : TDateTime ;
  function  SemaineSuivante(DateSem : TDateTime) : TDateTime;
  function  QuizaineSuivante(DateQuinz : TDateTime) : TDateTime;
  function  DateCloture(LastCloture : TDateTime) : TDateTime;
  function  RechDateFinExo : TDateTime;

  //Fonctions liées à la barre InitMove
  procedure ConstruitTobDepot(CDepot : THValComboBox);

  //fonctions de contrôles
  function  ListeInventOK(DateDeb, DateFin : TDateTime ; Mess : TStringList ; var TypEve : string) : boolean;//existence de listes d'inventaires sur la période et non validées
  function  ExistePieceVivante(DateFin : TDateTime) : boolean; //Existence de pièces vivantes

  //Procédures liées au journal d'événements
  Procedure NoteEvenement(Mess : TStringList ; Etat, LibEvent : string);
  procedure InitMessage(Mess : TStringList ; DateDeb, DateFin, Titre : string);


var StopTT : boolean; //True si annulation du traitement par l'utilisateur
    TobDepot : TOB;
    DateDebut : TDateTime;


Const
	// libellés des messages
	TexteMessage: array[1..9] of string 	= (
          {1}         'Le traitement n''a pu être effectué !',
          {2}         'Problèmes rencontrés : l''opération a été annulée',
          {3}         'Le stock ne peut être clôturé au delà du ',
          {4}         'La date de clôture doit être postérieure à la date de début de traitement',
          {5}         'Vous devez réaliser un traitement de fin d''exercice',
          {6}         'Vous n''avez pas validé toutes les listes d''inventaire ! Si vous clôturez elles seront perdues.'#13'Voulez vous continuer ?',
          {7}         'L''exercice en cours n''est pas arrivé à terme',
          {8}         'Avant de procéder à ce traitement vous devez clôturer les mois correspondant à l''exercice en cours',
          {9}         'Il est impossible de clôturer car il existe des pièces en cours affectant le stock'
                    );


implementation

////////////////////////////////////////////////////////////////////////////////
//************************* Fin de Mois **************************************//
////////////////////////////////////////////////////////////////////////////////

procedure TOF_TTFinMois.OnArgument (stArgument : String ) ;
var stPeriode : string;
    nbmois : integer;
begin
Inherited;
nbmois := StrToInt(GetParamSoc('SO_GCNBMOISSTOCKOUVERTS'));
if nbmois <= 0 then nbmois := 1;
SetControlText('HNBMOISOUVERTS',IntToStr(nbmois));
InitDates;
stPeriode := GetParamSoc('SO_GCPERIODESTOCK');
SetControlText('HPERIODICITE',RechDom('GCPERIODESTOCK',stPeriode,false));
TobDepot := TOB.Create('',nil,-1);
ConstruitTobDepot(THValComboBox(GetControl('DEPOT')));
end;

Procedure TOF_TTFinMois.OnUpdate ;
var FinClot : TDateTime;
    Evenement : string;
    Erreur    : boolean;
    Mess : TStringList;
begin
Inherited ;
if PGIAsk('Vous vous apprêter à lancer un long traitement,'#13'voulez vous continuer?','Traitement de fin de mois') = mrNo then exit;
if not DatesOk then exit;

FinClot := FinDeMois(StrToDate(GetControlText('HFinTT')));
if ExistePieceVivante(FinClot) then  begin LastError:=9 ; LastErrorMsg:=TexteMessage[LastError]; exit; end;

Mess := TStringList.Create;
InitMessage(Mess,GetControlText('HDebTT'),GetControlText('HFinTT'),Ecran.Caption);

if not ListeInventOK(StrToDate(GetControlText('HDebTT')), FinClot, Mess, Evenement) then exit;

InitVar;
Erreur:=false;

Try
  BeginTrans ;
  if TransactionFinMoisOK then CommitTrans
  else raise Exception.create('') ;
Except
  Erreur := true;
  if not StopTT then
     begin
     LastError:=2 ; LastErrorMsg:=TexteMessage[LastError];
     AfficheError:=False;
     PGIBox(TexteMessage[LastError],TForm(Ecran).Caption);
     end;

  RollBack ;
end ;

if  StopTT then
    begin
    Mess.Add('Traitement interrompu par l''utilisateur');
    NoteEvenement(Mess,'INT','M');
    end else
        begin
        if Erreur then
           begin
           LastError:=1 ; LastErrorMsg:=TexteMessage[LastError];
           AfficheError:=False;
           PGIBox(TexteMessage[LastError],TForm(Ecran).Caption);
           Mess.Add('Erreur - Période non clôturée');
           NoteEvenement(Mess,'ERR','M');
           end else
           begin
           PGIInfo('Le traitement s''est correctement effectué','Traitement de fin de mois');
           Mess.Add('Mois clôturés : du '+GetControlText('HDebTT')+' au '+GetControlText('HFinTT'));
           NoteEvenement(Mess,Evenement,'M');
           SetParamSoc('SO_GCDATECLOTURESTOCK',FinClot);
           VH_GC.GCDateClotureStock := FinClot; // DBR Fiche 10329 pour qu'immédiatement le param société en mémoire soit a jour
           InitDates;
           end;
        end;
Mess.Free;
end;

procedure TOF_TTFinMois.OnClose ;
begin
Inherited;
TobDepot.Free;
end;

/////////////////////////////Transaction
function TOF_TTFinMois.TransactionFinMoisOK : boolean;
var FinClot, FinPeriode : TDateTime;
begin
Result := true;
FinClot := FinDeMois(StrToDate(GetControlText('HFinTT')));
FinPeriode := DateCloture(StrToDate(GetControlText('HDebTT')));
while FinPeriode <= FinClot do
   begin
   Entree_Traitement(FinPeriode,false);
   if StopTT = true then
      begin
      result:=false ;
      exit;
      end;
   FinPeriode :=  DateCloture(FinPeriode);
   end;
end;

/////////////////////////////Contrôles sur les dates

procedure TOF_TTFinMois.InitDates;
var DernierTT, DebTT, FinTT, DateFinExo : TDateTime;
    DateMax : TDateTime;
    nbmois  : integer;
begin
if not isValidDate(GetParamSoc('SO_GCDATECLOTURESTOCK')) then
        SetParamSoc('SO_GCDATECLOTURESTOCK',RechDerniereCloture);
nbmois := StrToInt(GetControlText('HNBMOISOUVERTS'));
DateMax := FinDeMois(PlusMois(V_PGI.DateEntree,-nbmois));
DateFinExo := RechDateFinExo;
if DateMax >= DateFinExo then DateMax := PlusMois(DateFinExo,-1);

if (not ExisteSQL('SELECT GQ_ARTICLE FROM DISPO WHERE GQ_CLOTURE="X"'))//pas de cloture précédente
   or (StrToDate(GetParamSoc('SO_GCDATECLOTURESTOCK')) < FinDeMois(PlusMois(VH^.Encours.Deb,-1))) then //dernière date de cloture erronée
   begin
   DernierTT := VH^.Encours.Deb;
   DebTT := DernierTT;
   end else
   begin
   DernierTT := StrToDate(GetParamSoc('SO_GCDATECLOTURESTOCK'));
   DebTT:=DebutDeMois(PlusMois(DernierTT,1));
   end;

if DebTT > DateMax then FinTT := FinDeMois(DebTT)
   else FinTT := DateMax;

//if (FinTT >= VH^.Encours.Fin) or (FinTT >= VH^.Suivant.Fin) then
//    FinTT := FinDeMois(PlusMois(VH^.Encours.Fin,-1));

if DebTT > FinTT then FinTT := FinDeMois(DebTT);

SetControlText('HDernierTT',DateToStr(DernierTT));
SetControlText('HDebTT',DateToStr(DebTT));
DateDebut:=DebTT;
SetControlText('HFinTT',DateToStr(FinTT));
end;


//Contrôle des dates
function TOF_TTFinMois.DatesOk : boolean;
var DateFin : TDateTime;
    DateMax : TDateTime;
    DateFinExo : TDateTime;
    DernierClot : TDateTime;
    nbmois  : integer;
begin
nbmois := StrToInt(GetControlText('HNBMOISOUVERTS'));
DernierClot := StrToDate(GetParamSoc('SO_GCDATECLOTURESTOCK'));
DateFin :=  FinDeMois(StrToDate(GetControlText('HFinTT')));
DateMax :=  FinDeMois(PlusMois(V_PGI.DateEntree,-nbmois));
Result := True;

DateFinExo:=RechDateFinExo;

SetControlText('HFinTT',DateToStr(FinDeMois(DateFin)));
if DateMax>=DateFinExo then DateMax:=PlusMois(DateFinExo,-1);

//Date début > Date fin
if StrToDate(GetControlText('HDebTT')) > DateFin then
   begin
   LastError:=4 ; LastErrorMsg:=TexteMessage[LastError] ;
   AfficheError:=False;
   PGIBox(TexteMessage[LastError],TForm(Ecran).Caption);
   Result := False;
   exit;
   end;

//Date de clôture > Date possible de fin de TT
//if (FinDeMois(DateDeb) <> VH^.Encours.Fin) and (FinDeMois(DateDeb) <> VH^.Suivant.Fin) then
if DateFin > DateMax then
   begin
   if DernierClot = PlusMois(DateFinExo,-1) then
      begin
      LastError:=5 ; LastErrorMsg:=TexteMessage[LastError];
      AfficheError:=False;
      PGIBox(TexteMessage[LastError],TForm(Ecran).Caption);
      Result := False;
      exit;
      end;
   if (DateFin >= DateFinExo) and (DateMax = PlusMois(DateFinExo,-1)) then
      begin
      LastError:=3 ; LastErrorMsg:=TexteMessage[LastError];
      AfficheError:=False;
      PGIBox(TexteMessage[LastError] + DateToStr(DateMax) + '.#13' +
        'Au delà un traitement de fin d''exercice sera nécessaire',TForm(Ecran).Caption);
      Result := False;
      exit;
      end;
   LastError:=3 ; LastErrorMsg:=TexteMessage[LastError] ;
   AfficheError:=False;
   PGIBox(TexteMessage[LastError] + DateToStr(DateMax) + '#13' +
      'Nombre de mois ouverts : ' + IntToStr(nbmois),TForm(Ecran).Caption);
   Result := False;
   end;
end;

////////////////////////////////////////////////////////////////////////////////
//************************* Fin d'Exercice ***********************************//
////////////////////////////////////////////////////////////////////////////////

procedure TOF_TTFinExercice.OnArgument (stArgument : String ) ;
begin
Inherited;
InitDates;
TobDepot := TOB.Create('',nil,-1);
ConstruitTobDepot(THValComboBox(GetControl('DEPOT')));
end;

Procedure TOF_TTFinExercice.OnUpdate ;
var DateFinExo : TDateTime;
    Erreur : boolean;
    Mess : TStringList;
    stDebTT, stFinTT, Evenement : string;
begin
Inherited ;
if PGIAsk('Ce traitement peut être long, voulez vous continuer?','Traitement de fin d''exercice') = mrNo then exit;
DateFinExo := RechDateFinExo;

if not ControlDate(DateFinExo) then exit;
Mess := TStringList.Create;
stDebTT := DateToStr(DebutDeMois(DateFinExo));
DateDebut:=DebutDeMois(DateFinExo);
if ExistePieceVivante(DateFinExo) then
   begin
   LastError:=9 ; LastErrorMsg:=TexteMessage[LastError];
   AfficheError:=False;
   PGIBox(TexteMessage[LastError],TForm(Ecran).Caption);
   exit;
   end;
stFinTT := GetControlText('FINEXOFIN');
InitMessage(Mess,stDebTT,stFinTT,Ecran.Caption);

if not ListeInventOK(DebutDeMois(DateFinExo), DateFinExo, Mess, Evenement) then exit;

InitVar;
Erreur:=false;

Try
  BeginTrans ;
  if TransactionFinExoOK then CommitTrans
  else raise ERangeError.create('');
Except
 Erreur := true;
 RollBack ;
end ;

if StopTT then
   begin
   Mess.Add('Traitement interrompu par l''utilisateur');
   NoteEvenement(Mess,'INT','E');
   end else
       begin
       if Erreur then
          begin
          LastError:=1 ; LastErrorMsg:=TexteMessage[LastError];
          AfficheError:=False;
          PGIBox(TexteMessage[LastError],TForm(Ecran).Caption);
          Mess.Add('Erreur - Fin d''exercice non effectuée');
          NoteEvenement(Mess,'ERR','E');
          end else
          begin
          PGIInfo('Le traitement s''est correctement effectué','Traitement de fin d''exercice');
          Mess.Add('Mois clôturés : du '+stDebTT +' au '+stFinTT);
          Mess.Add('Le stock initial a été mis à jour');
          NoteEvenement(Mess,Evenement,'E');
          SetParamSoc('SO_GCDATECLOTURESTOCK',DateFinExo);
          VH_GC.GCDateClotureStock := DateFinExo; // DBR Fiche 10329 pour qu'immédiatement le param société en mémoire soit a jour
          end;
       end;
Mess.Free;
InitDates;
end;

procedure TOF_TTFinExercice.OnClose ;
begin
Inherited;
TobDepot.Free;
end;

///////////////////////Transaction
function TOF_TTFinExercice.TransactionFinExoOK : boolean;
begin
Result:=true;
Entree_Traitement(RechDateFinExo , true);
if StopTT = true then
   begin
   Result:=false;
   exit;
   end;
end;


//////////////////////Contrôle des dates
function TOF_TTFinExercice.ControlDate(DateFinExo : TDateTime) : boolean;
var DateClot, DateMax : TDateTime;
    nbmois : integer;
begin
Result := true;
DateClot := GetParamSoc('SO_GCDATECLOTURESTOCK');
if (V_PGI.DateEntree < DateFinExo) or (DateClot >= DateFinExo) then
   begin
   LastError:=7 ; LastErrorMsg:=TexteMessage[LastError];
   AfficheError:=False;
   PGIBox(TexteMessage[LastError],TForm(Ecran).Caption);
   Result:=false;
   exit;
   end;

if DateClot < PlusMois(DateFinExo,-1) then
   begin
   LastError:=8 ; LastErrorMsg:=TexteMessage[LastError];
   AfficheError:=False;
   PGIBox(TexteMessage[LastError],TForm(Ecran).Caption);
   Result:=false;
   exit;
   end;

nbmois := StrToInt(GetParamSoc('SO_GCNBMOISSTOCKOUVERTS'));
if nbmois <= 0 then nbmois := 1;
DateMax := FinDeMois(PlusMois(V_PGI.DateEntree,-nbmois));
if DateFinExo > DateMax then
   begin
   LastError:=3 ; LastErrorMsg:=TexteMessage[LastError] ;
   AfficheError:=False;
   PGIBox(TexteMessage[LastError] + DateToStr(DateMax) + '#13'
        + 'Nombre de mois ouverts : ' + IntToStr(nbmois),TForm(Ecran).Caption);
   Result:=false;
   exit;
   end;
end;

procedure TOF_TTFinExercice.InitDates;
var DateFinExo : TDateTime;
begin
if not isValidDate(GetParamSoc('SO_GCDATECLOTURESTOCK')) then
        SetParamSoc('SO_GCDATECLOTURESTOCK',RechDerniereCloture);
SetControlText('LASTCLOTURE',GetParamSoc('SO_GCDATECLOTURESTOCK'));
DateFinExo := RechDateFinExo;
if DateFinExo = VH^.Encours.Fin then
   begin
   SetControlText('FINEXODEB',DateToStr(VH^.Encours.Deb));
   SetControlText('FINEXOFIN',DateToStr(VH^.Encours.Fin));
   SetControlText('DEBEXOSUIV',DateToStr(VH^.Suivant.Deb));
   SetControlText('FINEXOSUIV',DateToStr(VH^.Suivant.Fin))
   end
   else if DateFinExo = VH^.Suivant.Fin then
   begin
   SetControlText('FINEXODEB',DateToStr(VH^.Suivant.Deb));
   SetControlText('FINEXOFIN',DateToStr(VH^.Suivant.Fin));
   SetControlText('DEBEXOSUIV','');
   SetControlText('FINEXOSUIV','');
   SetControlVisible('TDEBEXOSUIV',false);
   SetControlVisible('TFINEXOSUIV',false);
   end;
end;

////////////////////////////////////////////////////////////////////////////////
//**************** Procédures et fonction communes ***************************//
////////////////////////////////////////////////////////////////////////////////

/////////////////////////////Initialisations
procedure InitMessage(Mess : TStringList ; DateDeb, DateFin, Titre : string);
begin                                                                                  
Mess.Add(Titre);
Mess.Add('   Date de début : '+ DateDeb);
Mess.Add('   Date de fin : '+ DateFin);
end;

procedure InitVar ;
begin
  StopTT := False;
end;

function RechDerniereCloture : TDateTime;
var QLastDate : TQuery;
begin
result := 0;
QLastDate := OpenSQL('SELECT MAX(GQ_DATECLOTURE) AS DATECLOTURE FROM DISPO WHERE GQ_CLOTURE="X"', true);
if not QLastDate.Eof then Result := QLastDate.FindField('DATECLOTURE').AsDateTime;
Ferme(QLastDate);
if result <= 0 then result := 02;
end;

function DateCloture(LastCloture : TDateTime) : TDateTime;
begin
Result := LastCloture;
if GetParamSoc('SO_GCPERIODESTOCK')='SEM' then
   begin
   Result :=  SemaineSuivante(LastCloture);
   exit;
   end;
if GetParamSoc('SO_GCPERIODESTOCK')='QUI' then
   begin
   Result :=  QuizaineSuivante(LastCloture);
   exit;
   end;
if GetParamSoc('SO_GCPERIODESTOCK')='MOI' then
   begin
   if Copy(DateToStr(LastCloture),0,2) = '01' then
   Result :=  FinDeMois(LastCloture)
   else Result :=  FinDeMois(PlusMois(LastCloture,1));
   exit;
   end;
end;

function  RechDateFinExo : TDateTime;
begin
Result := VH^.Encours.Fin;
if StrToDate(GetParamSoc('SO_GCDATECLOTURESTOCK')) < VH^.Encours.Fin then
   Result:=VH^.Encours.Fin
   else if StrToDate(GetParamSoc('SO_GCDATECLOTURESTOCK')) < VH^.Suivant.Fin then
        Result:=VH^.Suivant.Fin;
//if StrToDate(GetParamSoc('SO_GCDATECLOTURESTOCK')) < FinDeMois(PlusMois(VH^.Encours.Deb,-1)) then
//   Result := FinDeMois(PlusMois(VH^.Encours.Deb,-1));
end;

function SemaineSuivante(DateSem : TDateTime) : TDateTime;
var jour : integer;
    stDateSem : string;
begin
Result := FinDeMois(DateSem);
stDateSem := DateToStr(DateSem);
jour := StrToInt(Copy(stDateSem,0,2));

case jour of
     01 : Result := StrToDate('08'+Copy(stDateSem,3,8)); //Début de TT
     08 : Result := StrToDate('15'+Copy(stDateSem,3,8));
     15 : Result := StrToDate('21'+Copy(stDateSem,3,8));
     21 : Result := FinDeMois(DateSem);
     else if jour > 21 then
          begin
          //Result := DebutDeMois(PlusMois(DateSem,1));
          stDateSem := DateToStr(PlusMois(DateSem,1));
          Result := StrToDate('08'+Copy(stDateSem,3,8));
          end;
     end;
end;


function QuizaineSuivante(DateQuinz : TDateTime) : TDateTime;
var jour : integer;
    stDateQuinz : string;
begin
Result := FinDeMois(DateQuinz);
stDateQuinz := DateToStr(DateQuinz);
jour := StrToInt(Copy(stDateQuinz,0,2));

case jour of
     01 : Result := StrToDate('15'+Copy(stDateQuinz,3,8));  //Début de TT
     15 : Result := FinDeMois(DateQuinz);
     else if jour > 15 then
          begin
          stDateQuinz := DateToStr(PlusMois(DateQuinz,1));
          Result := StrToDate('15'+Copy(stDateQuinz,3,8));
          end;
     end;
end;

/////////////////////////////Contrôle sur listes d'inventaire

function ListeInventOK(DateDeb, DateFin : TDateTime ; Mess : TStringList ; var TypEve : string) : boolean;
var QSelect : TQuery;
    CodesInvent : string;
begin
CodesInvent := '';
TypEve := 'OK';
Result := true;
QSelect := OpenSQL('SELECT GIE_CODELISTE FROM LISTEINVENT WHERE GIE_DATEINVENTAIRE >="'
                     + USDateTime(DateDeb) + '" AND GIE_DATEINVENTAIRE <="'
                     + USDateTime(DateFin) + '" AND GIE_VALIDATION="-"', false);

if not QSelect.Eof then
   if PGIAsk('Vous n''avez pas validé toutes les listes d''inventaire ! '
         + 'Si vous réalisez ce traitement elles seront perdues.'#13' Voulez vous continuer ?'
         , 'Traitement de fin d''exercice') = mrno then begin Ferme(QSelect) ; Result := false ; exit ;
   end else
   begin
   while not QSelect.Eof do
         begin
         CodesInvent := CodesInvent + QSelect.FindField('GIE_CODELISTE').AsString + ' ; ';
         QSelect.Next;
         end;
   Mess.Add('Listes d''inventaire perdues : ' + CodesInvent);
   TypEve := 'AVE';
   end;
   Ferme(QSelect);
end;

function  ExistePieceVivante(DateFin : TDateTime) : boolean;
var QPiece : TQuery;
    stWhere : string;
begin
Result:=false;
stWhere := RecupWhereNaturePiece('PHY','GP');
   if stWhere <> '' then stWhere := ' AND ' + stWhere;
QPiece := OpenSQL('SELECT COUNT(GP_NUMERO) AS NBPIECE '+
             'FROM PIECE WHERE GP_DATEPIECE>="'+USDateTime(DateDebut)+'" '+
             'AND GP_DATEPIECE<="'+USDateTime(DateFin)+'" '+
             'AND GP_VIVANTE="X"' + stWhere ,true);
if (not QPiece.Eof) and (QPiece.FindField('NBPIECE').AsInteger > 0) then Result:=true;
Ferme(QPiece);
end;

/////////////////////////////Insertion dans le journal d'événements

Procedure NoteEvenement(Mess : TStringList ; Etat, LibEvent : string);
var TobJNL : TOB;
    Indice : integer;
    QIndice : TQuery;
begin
Indice := 0;
QIndice := OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT',true);
if not QIndice.Eof then
     Indice := QIndice.Fields[0].AsInteger + 1;
Ferme(QIndice);

TobJNL := TOB.Create('JNALEVENT',nil,-1);

TobJNL.PutValue('GEV_NUMEVENT',Indice);
TobJNL.PutValue('GEV_TYPEEVENT','STK');
if LibEvent = 'M' then TobJNL.PutValue('GEV_LIBELLE','Fin de mois')
                  else TobJNL.PutValue('GEV_LIBELLE','Fin d''exercice');
TobJNL.PutValue('GEV_DATEEVENT',V_PGI.DateEntree);
TobJNL.PutValue('GEV_UTILISATEUR',V_PGI.USer);
TobJNL.PutValue('GEV_ETATEVENT',Etat);
TobJNL.PutValue('GEV_BLOCNOTE', Mess.Text);
TobJNL.InsertOrUpdateDB(false);
end;


/////////////////////////////Préparation des données dans la TOBDispo

procedure Entree_Traitement(DatClot : TDateTime ; FinExo : boolean);
var i_dep,i_art : integer;
    st,TitreTQR, stRequete, stOrder, stDepot : string;
    TobDispoArt,TobDispo : TOB;
    QQ : TQuery ;
    NbEnreg : integer;
begin
//Initialisation
NbEnreg:=0;
QQ:=OpenSQL('SELECT Count(GQ_ARTICLE) as nombre FROM DISPO WHERE GQ_CLOTURE="-" AND GQ_ARTICLE<>"" ',True) ;
if not QQ.Eof then
    begin
    NbEnreg:=QQ.FindField ('nombre').AsInteger;
    end;
Ferme(QQ);
if FinExo then st:='Fin d''exercice' else st:='Fin de mois';
InitMoveProgressForm (nil,st,'Traitement en cours...', NbEnreg, true, true) ;
if FinExo then TitreTQR:='Date de fin d''exercice : ' + DateToStr(DatClot)
else TitreTQR:='Date de clôture : ' + DateToStr(DatClot);
stRequete := 'SELECT GQ_ARTICLE FROM DISPO WHERE GQ_CLOTURE="-" AND GQ_ARTICLE<>"" '
                 + 'AND GQ_DEPOT=';
stOrder := ' ORDER BY GQ_ARTICLE';
TobDispoArt := TOB.Create('',nil,-1);

/////Traitement
for i_dep := 0 to TobDepot.Detail.Count -1 do
    begin
    stDepot := TobDepot.Detail[i_dep].GetValue('DEPOT');
    TobDispoArt.LoadDetailFromSQL(stRequete+'"'+stDepot+'"'+stOrder);
    for i_art := 0 to TobDispoArt.Detail.Count -1 do
       begin
       MoveCur(False);
       if not MoveCurProgressForm(TitreTQR+'   Dépôt : '+TobDepot.Detail[i_dep].GetValue('LIBELLE')) then
          begin
          StopTT := True;
          Break;
          end;
       TobDispo := TOB.Create('DISPO',nil,-1);  TobDispo.InitValeurs;
       TobDispo.PutValue('GQ_ARTICLE',TobDispoArt.Detail[i_art].GetValue('GQ_ARTICLE'));
       TobDispo.PutValue('GQ_DEPOT',stDepot);
       MAJQuantite (TobDispo,DatClot,FinExo);
       ClotureMois(TobDispo,DatClot);
       TobDispo.SetAllModifie(true);
       TobDispo.InsertDB(nil,false);
       TobDispo.Free;
       end;
    TobDispoArt.ClearDetail;
    if StopTT then Break;
    end;
TobDispoArt.Free;
FiniMoveProgressForm ;
end;

procedure MAJQuantite (TOBTT : TOB ; DatClot : TDateTime ; FinExo : boolean);
var  Article,Depot : string;
     TQte : TQtePrixRec;
begin
Article := TOBTT.GetValue('GQ_ARTICLE');  Depot := TOBTT.GetValue('GQ_DEPOT');
TQte := TQtePrixRec(GetQtePrixDateListe(Article,Depot,DatClot));

if TQte.SomethingReturned then
   begin
   TOBTT.PutValue('GQ_PHYSIQUE',Arrondi(TQte.Qte,V_PGI.OkDecQ));
   TOBTT.PutValue('GQ_CUMULENTREES',TQte.QtePlus);
   TOBTT.PutValue('GQ_CUMULSORTIES',TQte.QteMoins);
   TOBTT.PutValue('GQ_DPA',TQte.DPA);
   TOBTT.PutValue('GQ_DPR',TQte.DPR);
   TOBTT.PutValue('GQ_PMAP',TQte.PMAP);
   TOBTT.PutValue('GQ_PMRP',TQte.PMRP);
   end else
   begin //pas de clôture et pas de pièce
   TOBTT.PutValue('GQ_PHYSIQUE',0);
   TOBTT.PutValue('GQ_CUMULENTREES',0);
   TOBTT.PutValue('GQ_CUMULSORTIES',0);
   TOBTT.PutValue('GQ_DPA',0);
   TOBTT.PutValue('GQ_DPR',0);
   TOBTT.PutValue('GQ_PMAP',0);
   TOBTT.PutValue('GQ_PMRP',0);
   end;

if FinExo then  ExecuteSQL('UPDATE DISPO SET GQ_STOCKINITIAL="'+StrfPoint(TQte.Qte) +
                           ', GQ_STOCKINV="'+StrfPoint(TQte.Qte) +
                           '" WHERE GQ_ARTICLE="'+Article+
                           '" AND GQ_DEPOT="'+Depot+'" AND GQ_CLOTURE="-"');

TOBTT.PutValue('GQ_LIVRECLIENT',0);
TOBTT.PutValue('GQ_LIVREFOU',0);
TOBTT.PutValue('GQ_RESERVECLI',0);
TOBTT.PutValue('GQ_RESERVEFOU',0);
TOBTT.PutValue('GQ_QTE1',0);
TOBTT.PutValue('GQ_QTE2',0);
TOBTT.PutValue('GQ_QTE3',0);
end;

procedure ClotureMois(TOBTT : TOB ; DatClot : TDateTime);
begin
TOBTT.PutValue('GQ_CLOTURE','X');
TOBTT.PutValue('GQ_DATECLOTURE',DatClot);
TOBTT.PutValue('GQ_DATECREATION',V_PGI.DateEntree);
end;


////////////////////////////////////////////////////////////////////////////////
//******************* Fonctions liées à la barre InitMove ********************//
////////////////////////////////////////////////////////////////////////////////

procedure ConstruitTobDepot(CDepot : THValComboBox);
var i_ind : integer;
    TobD : TOB;
begin
for i_ind := 0 to CDepot.Items.Count -1 do
    begin
    TobD := TOB.Create('',TobDepot,-1);
    TobD.AddChampSup('DEPOT',false); TobD.AddChampSup('LIBELLE',false);
    TobD.PutValue('DEPOT',CDepot.Values[i_ind]);
    TobD.PutValue('LIBELLE',CDepot.Items[i_ind]);
    end;
end;

Initialization
registerclasses([TOF_TTFinMois]);
registerclasses([TOF_TTFinExercice]);
end.

