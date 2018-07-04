unit ImoUtil;

interface
uses dbtables,HCtrls,Hent1, S1Util,Utob ;

Const MaxAxe  = 5 ;

Type TExoDate = record
  Code    : String[3] ;
  Deb,Fin : TDateTime ;
  end ;

  TDefCompte = record
    Compte:string;
    Libelle:string;
    bGeneral : boolean;
  end;

  TInfoCpta = record
    Lg : Byte ;
    Cb : Char ;
  end ;

{ne pas changer l'ordre des 5 premiers (axes)}
Type TFichierBase = (fbAxe1,fbAxe2,fbAxe3,fbAxe4,fbAxe5,fbGene,fbAux,fbJal,fbNone) ;


//-Imo
type LaVariableImo = record
     TenueEuro:  boolean ;
     EnCours,Suivant : TExoDate ;
     Cpta : Array[fbAxe1..fbAux] of TInfoCpta ;
     OBImmo   : TOB ;
   // Immobilisations
     Exercices : Array [1..20] of TExoDate ;
     CpteCBInf,CpteCBSup : string[17];
     CpteLocInf,CpteLocSup : string[17];
     CpteDepotInf,CpteDepotSup : string[17];
     CpteImmoInf,CpteImmoSup : string[17];
     CpteFinInf,CpteFinSup : string[17];
     CpteAmortInf,CpteAmortSup : string[17];
     CpteDotInf,CpteDotSup : string[17];
     CpteExploitInf,CpteExploitSup : string[17];
     CpteDotExcInf,CpteDotExcSup : string[17];
     CpteRepExcInf,CpteRepExcSup : string[17];
     CpteVaCedeeInf,CpteVaCedeeSup : string[17];
     CpteDerogInf,CpteDerogSup : string[17];
     CpteProvDerInf,CpteProvDerSup : string[17];
     CpteRepDerInf,CpteRepDerSup : string[17];
     Specif : Array[0..9] Of Boolean ;
     ChargeOBImmo : boolean;
     SpeedCum : Boolean ;
end ;

var VH: ^LaVariableImo ;
implementation
procedure InitLaVariableIMO ;
begin
  New(VH) ;
  FillChar(VH^,Sizeof(VH^),#0) ;
end ;


procedure ChargeIMO ;
var Premier: boolean ; DDeb,DFin : TDateTime ;  SCode,SEtat : String ; i: integer ; Q: TQuery ;
begin
  VH^.TenueEuro:=VS1^.TenueEuro  ;
  i:=1 ; Premier:=true ;
  Q:=OpenSQL('SELECT * FROM EXERCICE ORDER BY EX_DATEDEBUT',TRUE) ;
  try
    if not Q.Eof then
    begin
      Q.First ;
      while ((not Q.Eof) and (i<=20)) do
      begin
        DDeb:=Q.FindField('EX_DATEDEBUT').AsDateTime ;
        DFin:=Q.FindField('EX_DATEFIN').AsDateTime ;
        SCode:=Q.FindField('EX_EXERCICE').AsString ;
        SEtat:=Q.FindField('EX_ETATCPTA').AsString ;
        VH^.Exercices[i].Code := SCode ;
        VH^.Exercices[i].Deb := DDeb ;
        VH^.Exercices[i].Fin := DFin ;
        if (SEtat='OUV') Or (SEtat='CPR') then
        begin
          if Premier then
          begin
            VH^.Encours.Deb:=DDeb ;
            VH^.Encours.Fin:=DFin ;
            VH^.Encours.Code:=SCode ;
          end
          else
          begin
            VH^.Suivant.Deb:=DDeb ;
            VH^.Suivant.Fin:=DFin ;
            VH^.Suivant.Code:=SCode ;
            Break ;
          end ;
          Premier:=False ;
        end ;
        Inc(i) ;
        Q.Next;
      end;
    end
  finally
    Ferme(Q) ;
    VH^.Exercices[i].Code := '';
    VH^.Exercices[i].Deb := iDate1900;
    VH^.Exercices[i].Fin := iDate1900;
  end ;
  //Charge CptaInfo
  VH^.Cpta[fbGene].Lg:=VS1^.LEN_GENE ;
  VH^.Cpta[fbGene].Cb:='0' ;
  VH^.Cpta[fbAux].Lg:=VS1^.LEN_GENE ;
  VH^.Cpta[fbAux].Cb:='0' ;
end ;



end.
