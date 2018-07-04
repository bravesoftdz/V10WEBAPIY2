unit BPMaille;

interface

uses HEnt1,classes;

type
  TMaille = class(TObject)
    code:integer;
    libelle:hString;
    DateDebCourante:TDateTime;
    DateFinCourante:TDateTime;
    DateDebReference:TDateTime;
    DateFinReference:TDateTime;
    types:integer;  //1 : jour
                    //2 : semaine
                    //3 : quinzaine
                    //4 : mois
                    //5 : mois 4-4-5
  end;

  TListMaille=class(TList);

procedure RemplitTMaille(code:integer;const libelle:hString;
          DateDebC,DateFinC,DateDebR,DateFinR:TDateTime;
          types:integer;
          resultat:TMaille);
procedure InitialiseListeMaille(types:integer;
          DateDebC,DateFinC,DateDebR,DateFinR:TDateTime;
          ListMaille:TListMaille);
function  DateDonneMaille(DateX:TDateTime;ListMaille:TListMaille):integer;
function  MailleDonneDate(Maille:integer;ListMaille:TListMaille):TDateTime;
function  MailleDonneDateDebRef(Maille:integer;ListMaille:TListMaille):TDateTime;

procedure freeListMaille(L: TListMaille);

implementation

uses sysutils,BPBasic;

procedure freeListMaille(L: TListMaille);
var i:integer;
begin
  for i:=0 to L.count-1 do
  begin
    TMaille(L[i]).free;
  end;
  L.free;
end;

procedure RemplitTMaille(code:integer;const libelle:hString;
          DateDebC,DateFinC,DateDebR,DateFinR:TDateTime;
          types:integer;
          resultat:TMaille);
begin
  Resultat.code:=code;
  Resultat.libelle:=libelle;
  Resultat.DateDebCourante:=DateDebC;
  Resultat.DateFinCourante:=DateFinC;
  Resultat.DateDebReference:=DateDebR;
  Resultat.DateFinReference:=DateFinR;
  Resultat.types:=types;
end;

procedure InitialiseListeMaille(types:integer;
          DateDebC,DateFinC,DateDebR,DateFinR:TDateTime;
          ListMaille:TListMaille);
var Maille:TMaille;
    codeMi:integer;
    libelleMi:hString;
    DateDebCMi,DateFinCMi,DateDebRMi,DateFinRMi:TDateTime;
    i:integer;
    DateCI,DateRI,DateQCI,DateQRI:TDateTime;
    yCI,mCI,dCI,yRI,mRI,dRI:word;
    per: TperiodeP;
    NumT,AnT,NumTF,AnTF,NsRI,NsCI,anCI,anRI:integer;
begin
 case types of
  1 : begin
       //jour
       i:=1;
       DateCI:=DateDebC;
       DateRI:=DateDebR;
       while DateCI<=DateFinC do
        begin
         DecodeDate(DateCI,yCI,mCI,dCI);
         DecodeDate(DateRI,yRI,mRI,dRI);
         DateDebCMi:=DateCI;
         DateFinCMi:=DateCI;
         DateDebRMi:=DateRI;
         DateFinRMi:=DateRI;
         codeMi:=i;
         libelleMi:='';

         DateCI:=PlusDate(DateDebCMi,1,'J');
         DateRI:=PlusDate(DateDebRMi,1,'J');
         i:=i+1;

         //ajout dans la liste de maille
         Maille:=TMaille.Create;
         RemplitTMaille(codeMi,libelleMi,DateDebCMi,DateFinCMi,
                        DateDebRMi,DateFinRMi,types,Maille);
         ListMaille.Add(Maille);
        end;
      end;
  2 : begin
       //semaine
       i:=1;
       DateCI:=DateDebC;
       DateRI:=DateDebR;
       while DateCI<DateFinC do
        begin
         NsCI:=numsemaine(DateCI,anCI);
         NsRI:=numsemaine(DateRI,anRI);
         DateDebCMi:=PremierJourSemaine(NsCI,anCI);
         DateFinCMi:=PlusDate(DateDebCMi,6,'J');
         DateDebRMi:=PremierJourSemaine(NsRI,anRI);
         DateFinRMi:=PlusDate(DateDebRMi,6,'J');
         codeMi:=i;
         libelleMi:='';
         //ajout dans la liste de maille
         Maille:=TMaille.Create;
         RemplitTMaille(codeMi,libelleMi,DateDebCMi,DateFinCMi,
                        DateDebRMi,DateFinRMi,types,Maille);
         ListMaille.Add(Maille);
         DateCI:=PlusDate(DateDebCMi,7,'J');
         DateRI:=PlusDate(DateDebRMi,7,'J');
         i:=i+1;
        end;
      end;
  3 : begin
       //quinzaine
       i:=1;
       DateCI:=DateDebC;
       DateRI:=DateDebR;
       while DateCI<DateFinC do
        begin
         DateDebCMi:=DEBUTDEMOIS(DateCI);
         DateFinCMi:=FINDEMOIS(DateCI);    
         DateDebRMi:=DEBUTDEMOIS(DateRI);
         DateFinRMi:=FINDEMOIS(DateRI);
         DateQCI:=PlusDate(DateDebCMi,14,'J');
         DateQRI:=PlusDate(DateDebRMi,14,'J');
         codeMi:=i;
         libelleMi:='';

         DateCI:=PLUSMOIS(DateCI,1);
         DateRI:=PLUSMOIS(DateRI,1);
         i:=i+1;

         //ajout dans la liste de maille
         //1° quinzaine
         Maille:=TMaille.Create;
         RemplitTMaille(codeMi,libelleMi,DateDebCMi,DateQCI,
                        DateDebRMi,DateQRI,types,Maille);
         ListMaille.Add(Maille);
         //2° quinzaine
         { EVI / Bug : une quinzaine supplémentaire était ajouté si la session se terminait un 15 }
         if DateQCi <> DateFinC then
         begin
           Maille:=TMaille.Create;
           RemplitTMaille(codeMi,libelleMi,PlusDate(DateQCI,1,'J'),DateFinCMi,
                          PlusDate(DateQRI,1,'J'),DateFinRMi,types,Maille);
           ListMaille.Add(Maille);
         end;
        end;
      end;
  4 : begin
       //mois
       i:=1;
       DateCI:=DateDebC;
       DateRI:=DateDebR;
       while DateCI<DateFinC do
        begin
         DateDebCMi:=DEBUTDEMOIS(DateCI);
         DateFinCMi:=FINDEMOIS(DateCI);
         DateDebRMi:=DEBUTDEMOIS(DateRI);
         DateFinRMi:=FINDEMOIS(DateRI);
         codeMi:=i;
         libelleMi:='';

         DateCI:=PLUSMOIS(DateCI,1);
         DateRI:=PLUSMOIS(DateRI,1);
         i:=i+1;

         //ajout dans la liste de maille
         Maille:=TMaille.Create;
         RemplitTMaille(codeMi,libelleMi,DateDebCMi,DateFinCMi,
                        DateDebRMi,DateFinRMi,types,Maille);
         ListMaille.Add(Maille);
        end;
      end;
  5 : begin
       //mois 4-4-5
       i:=1;
       DateCI:=DateDebC;
       DateRI:=DateDebR;
       while DateCI<DateFinC do
        begin
         codeMi:=i;
         libelleMi:='';

         Per:=FindPeriodeP(DateRI);
         DateDebRMi:=Per.datedeb;
         DateFinRMi:=Per.datefin;
         Per:=FindPeriodeP(DateCI);
         DateDebCMi:=Per.datedeb;
         DateFinCMi:=Per.datefin;

         //ajout dans la liste de maille
         Maille:=TMaille.Create;
         RemplitTMaille(codeMi,libelleMi,DateDebCMi,DateFinCMi,
                        DateDebRMi,DateFinRMi,types,Maille);
         ListMaille.Add(Maille);


         Per:=AddPeriodeP(DateCI,1);
         DateCI:=Per.datedeb;
         Per:=AddPeriodeP(DateRI,1);
         DateRI:=Per.datedeb;     
         i:=i+1;
        end;
      end;  
  6 : begin
       //trimestre
       i:=1;
       DateCI:=DateDebC;
       DateRI:=DateDebR;
       while DateCI<DateFinC do
        begin
         codeMi:=i;
         libelleMi:='';

         DonneNumTrimestre(DateCI,NumT,AnT);
         DonneDateDebFinTrimestre(NumT,AnT,DateDebCMi,DateFinCMi);
         DonneNumTrimestre(DateRI,NumTF,AnTF);
         DonneDateDebFinTrimestre(NumTF,AnTF,DateDebRMi,DateFinRMi);
         //ajout dans la liste de maille
         Maille:=TMaille.Create;
         RemplitTMaille(codeMi,libelleMi,DateDebCMi,DateFinCMi,
                        DateDebRMi,DateFinRMi,types,Maille);
         ListMaille.Add(Maille);

         DateCI:=PLUSMOIS(DateCI,3);
         DateRI:=PLUSMOIS(DateRI,3);

         i:=i+1;
        end;
      end;
  7 : begin
       //quadrimestre
       i:=1;
       DateCI:=DateDebC;
       DateRI:=DateDebR;
       while DateCI<DateFinC do
        begin
         codeMi:=i;
         libelleMi:='';

         DonneNumQuadrimestre(DateCI,NumT,AnT);     
         DonneDateDebFinQuadrimestre(NumT,AnT,DateDebCMi,DateFinCMi);
         DonneNumQuadrimestre(DateRI,NumTF,AnTF);
         DonneDateDebFinQuadrimestre(NumTF,AnTF,DateDebRMi,DateFinRMi);
         //ajout dans la liste de maille
         Maille:=TMaille.Create;
         RemplitTMaille(codeMi,libelleMi,DateDebCMi,DateFinCMi,
                        DateDebRMi,DateFinRMi,types,Maille);
         ListMaille.Add(Maille);

         DateCI:=PLUSMOIS(DateCI,4);
         DateRI:=PLUSMOIS(DateRI,4);
         i:=i+1;
        end;
      end;
   else
    begin
     
     //ajout dans la liste de maille
     Maille:=TMaille.Create;
     RemplitTMaille(1,'1',DateDebC,DateFinC,
                    DateDebR,DateFinR,types,Maille);
     ListMaille.Add(Maille);
    end;
  end;
end;

function DateDonneMaille(DateX:TDateTime;ListMaille:TListMaille):integer;
var i:integer;
    cont:boolean;
    MailleMi:TMaille;
begin
 result:=0;
 cont:=true;
 i:=1;
 while (cont) and (i<ListMaille.count) do
  begin
   MailleMi:=TMaille(ListMaille[i-1]);
   if (DateX>=MailleMi.DateDebCourante) and (DateX<MailleMi.DateFinCourante)
    then
     begin
      cont:=false;
      result:=i;
     end;
   i:=i+1;
  end;
end;

function MailleDonneDate(Maille:integer;ListMaille:TListMaille):TDateTime;
var i:integer;
    cont:boolean;
    MailleMi:TMaille;
begin
 result:=0;
 cont:=true;
 i:=1;
 while (cont) and (i<=ListMaille.count) do
  begin
   MailleMi:=TMaille(ListMaille[i-1]);
   if (Maille=MailleMi.code)
    then
     begin
      cont:=false;
      result:=MailleMi.DateFinCourante;
     end;
   i:=i+1;
  end;
end;

function MailleDonneDateDebRef(Maille:integer;ListMaille:TListMaille):TDateTime;
var i:integer;
    cont:boolean;
    MailleMi:TMaille;
begin
 result:=0;
 cont:=true;
 i:=1;
 while (cont) and (i<=ListMaille.count) do
  begin
   MailleMi:=TMaille(ListMaille[i-1]);
   if (Maille=MailleMi.code)
    then
     begin
      cont:=false;
      result:=MailleMi.DateDebReference;
     end;
   i:=i+1;
  end;
end;

end.
