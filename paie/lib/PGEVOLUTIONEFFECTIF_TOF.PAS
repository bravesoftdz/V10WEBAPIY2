{***********UNITE*************************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 26/02/2004        
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGEVOLUTIONEFFECTIF ()
Mots clefs ... : TOF;PGEVOLUTIONEFFECTIF
*****************************************************************}
{
 PT1 : 07/10/2005 MF V610 FQ 12614 :
                            1-le traitement est effectué sur la période
                            choisie et non plus sur l'exercice complet.
                            2-Le calcul des effectifs prend en compte soit tous
                            les salariés présents au cours du mois soit uniquement
                            ceux présents à fin de période (Boîte à cocher)
PT2 : 12/10/2005 JL FQ 12222 Les efefctifs n'etaient pas additionné
                             + modif affectation tobeff
PT3 : 25/09/2009 SJ FQ n°12376 Suppression du LanceEtatTob
}

Unit PGEVOLUTIONEFFECTIF_TOF ;

Interface

Uses
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     QRS1,//PT3
     EdtREtat,
{$else}
     UtileAGL,
    eQRS1,//PT3
{$ENDIF}
     StdCtrls,     // PT1-2              
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HQry,
     ParamDat,
     ParamSoc,
     UTob,
     HTB97,
     UTOF ;

Type
  TOF_PGEVOLUTIONEFFECTIF = Class (TOF)
  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure OnUpdate                 ; override ;//PT3

  private
    ExDateDeb, ExDateFin, Etab,Categ,Annee      : string;
    TobEff, TobEtat,T                           : TOB;

    procedure BEtatClick(Sender : TObject);
    procedure DateElipsisclick(Sender: TObject);
    procedure ChangeExercice(Sender: TObject);
    procedure ControlPeriode(Sender: TObject);
    procedure ConstruireTob ();
    procedure PgIFValidPeriode;
  end ;

Implementation

uses EntPaie;

procedure TOF_PGEVOLUTIONEFFECTIF.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGEVOLUTIONEFFECTIF.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGEVOLUTIONEFFECTIF.BEtatClick(Sender : TObject);
var
  Where,GroupBy,SQL             : String;
  Pages                         : TPageControl;
  TobEffCat                     : TOB;
  QEff                          : TQuery;
  i                             : integer;
  FinPeriode                    : TCheckBox;                 // PT1-2
  begin
{ TobEtat := NIL;} //PT3 mise en commentaire
  If TobEtat <> Nil Then FreeAndNil(TobEtat); //PT3
  Pages := TPageControl(GetControl('Pages'));
// d PT1-2
  Where := '';
  Where := RecupWhereCritere(Pages);
  if (Where <> '') then
     Where := Where + ' AND PHC_CUMULPAIE="01" AND PSA_DADSPROF <> "" '
  else
     Where := 'WHERE PHC_CUMULPAIE="01" AND PSA_DADSPROF <> "" ';
//PT1                 ' AND YEAR(PHC_DATEFIN)="'+Annee+'" ';

  FinPeriode := TCheckBox(GetControl('FINPERIODE'));
  if (FinPeriode <> nil) then
    if FinPeriode.Checked = true then
      Where :=  Where + ' AND (PSA_DATESORTIE <= "'+UsDateTime(IDate1900)+'"'+
                        ' OR  PSA_DATESORTIE IS NULL '+
                        ' OR PSA_DATESORTIE > "'+
                        UsDateTime(StrtoDate(GetControlText('PHC_DATEFIN')))+'")';

//  Where := Where + RecupWhereCritere(Pages);
// f PT1-2

  SQL := 'SELECT PHC_ETABLISSEMENT,ET_LIBELLE,PSA_DADSPROF CATEGORIE,'+
         'COUNT (DISTINCT (PHC_SALARIE)) EFF, MONTH(PHC_DATEFIN) MOIS '+
         'FROM HISTOCUMSAL '+
         'lEFT JOIN ETABLISS ON ET_ETABLISSEMENT=PHC_ETABLISSEMENT '+
         'LEFT JOIN SALARIES ON PSA_SALARIE=PHC_SALARIE ';
  GroupBy := 'GROUP BY PHC_ETABLISSEMENT,ET_LIBELLE,PSA_DADSPROF,MONTH(PHC_DATEFIN)';
  SQL := SQL + Where + GroupBy;
  QEff := OpenSql(SQL,TRUE);
  TobEffCat := Tob.Create('Les effectifs',Nil,-1);
  TobEffCat.LoadDetailDB('Les effectifs','','',QEff,False);
  Ferme(QEff);

  if (TOBEtat = NIL) then
  begin
    TOBEtat := TOB.Create('Effectifs à éditer', NIL, -1);
  end;

  Etab := '';
  Categ := '';
//  TobEff := TobEffCat.FindFirst([''],[''],TRUE);

  for i := 0 to TobEffCat.Detail.Count -1 do
  begin
    TobEff := TobEffCat.Detail[i];     //PT2
    ConstruireTob;
//    TobEff := TobEffCat.FindNext([''],[''],TRUE);
  end;
  TobEffCat.free;

{if V_PGI.Debug=True then
   PGVisuUnObjet( TobEtat,'','') ;}

//debut PT3
  {LanceEtatTOB('E','PEV','PEM',TobEtat,True,False,False,Pages,' ','',False);
  TobEtat.Free;}//mise en commentaire
  TFQRS1(Ecran).LaTob:= TobEtat;
//fin PT3
end ;

procedure TOF_PGEVOLUTIONEFFECTIF.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGEVOLUTIONEFFECTIF.OnArgument (S : String ) ;
var
  THDateDeb,THDateFin           : THEdit;
  Exercice                      : THValComboBox ;
  QPeriode,QExer                : TQuery;
  {BEtat : TToolBarButton97;}//PT3 mise en commentaire 

begin
  Inherited ;
  SetControlText('DOSSIER',GetParamSocSecur ('SO_LIBELLE',''));
  SetControlChecked('PAIEDECALEE',VH_Paie.PGDecalage);

  QExer:=OpenSql('SELECT MAX(PEX_EXERCICE),MAX(PEX_DATEDEBUT),MAX(PEX_DATEFIN)'+
                 ',MAX(PEX_ANNEEREFER)'+
                 ' FROM EXERSOCIAL '+
                 'WHERE PEX_ACTIF="X"',true);
  if not QExer.eof then
  begin
    SetControlText('EDEXERSOC',QExer.Fields[0].asstring);
    ExDateDeb :=QExer.Fields[1].asstring;
    ExDateFin :=QExer.Fields[2].asstring;
    Annee := QExer.Fields[3].asstring;
  end
  else
  begin
    SetControlText('EDEXERSOC','');
    ExDateDeb:=DateToStr(idate1900);
    ExDateFin:=DateToStr(idate1900);
    Annee := '';
  end;
  Ferme(QExer);

// d PT1-1
//  THDateDeb:= ThEdit(getcontrol('XX_VARIABLEDEB'));
//  THDateFin:= ThEdit(getcontrol('XX_VARIABLEFIN'));
  THDateDeb:= ThEdit(getcontrol('PHC_DATEDEBUT'));
  THDateFin:= ThEdit(getcontrol('PHC_DATEFIN'));
// f PT1-1

  QPeriode:=OpenSql('SELECT MIN(PPU_DATEDEBUT),MAX(PPU_DATEFIN) '+
                    'FROM PAIEENCOURS '+
                    'LEFT JOIN EXERSOCIAL ON PPU_DATEDEBUT>=PEX_DATEDEBUT AND '+
                    'PPU_DATEFIN<=PEX_DATEFIN '+
                    'WHERE PEX_EXERCICE="'+GetControltext('EDEXERSOC')+'"',TRUE);
  if not QPeriode.eof then
  begin
    ExDateDeb:=DateToStr(QPeriode.Fields[0].AsDateTime);
    ExDateFin:=DateToStr(QPeriode.Fields[1].AsDateTime);
  end;
  Ferme(QPeriode);

  Exercice:= THValComboBox(getcontrol('EDEXERSOC'));
  if (THDateDeb <> nil) and (THDateFin <> nil) and (Exercice<>nil) then
  begin
   THDateDeb.text := ExDateDeb;
   THDateDeb.OnElipsisClick := DateElipsisclick;
   THDateDeb.OnExit := ControlPeriode;
   THDateFin.text := ExDateFin;
   THDateFin.OnElipsisClick := DateElipsisclick;
   THDateFin.OnExit := ControlPeriode;
   Exercice.OnChange := ChangeExercice;
  end;

  {BEtat  := TToolBarButton97(GetControl('BVALIDER'));
  if Betat <> Nil then BEtat.OnClick  := BEtatClick;}////PT3 mise en commentaire
end ;

procedure TOF_PGEVOLUTIONEFFECTIF.OnClose ;
begin
  Inherited ;
 If TobEtat <> Nil Then FreeAndNil(TobEtat)//PT3
end ;

//debut PT3
procedure TOF_PGEVOLUTIONEFFECTIF.OnUpdate;
begin
BEtatClick(nil);
end;
//fin PT3

procedure TOF_PGEVOLUTIONEFFECTIF.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGEVOLUTIONEFFECTIF.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_PGEVOLUTIONEFFECTIF.DateElipsisclick(Sender: TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGEVOLUTIONEFFECTIF.PgIFValidPeriode;
var
  YYD,MMD,JJ,YYF,MMF : WORD;
begin
// d PT1-1
//  if IsValidDate(GetControlText('XX_VARIABLEDEB')) AND
//   IsValidDate(GetControlText('XX_VARIABLEFIN')) Then
  if IsValidDate(GetControlText('PHC_DATEDEBUT')) AND
   IsValidDate(GetControlText('PHC_DATEFIN')) Then
  begin
//    DecodeDate(StrToDate(GetControlText('XX_VARIABLEDEB')),YYD,MMD,JJ);
//    DecodeDate(StrToDate(GetControlText('XX_VARIABLEFIN')),YYF,MMF,JJ);
    DecodeDate(StrToDate(GetControlText('PHC_DATEDEBUT')),YYD,MMD,JJ);
    DecodeDate(StrToDate(GetControlText('PHC_DATEFIN')),YYF,MMF,JJ);
    if (YYF>YYD) and (MMF>=MMD) then
    begin
      PgiBox('La période d''édition ne peut excéder douze mois civils.','Date Erronée!');
//      SetControlText('XX_VARIABLEFIN',DateToStr(FinDeMois(PlusDate(StrToDate(GetControlText('XX_VARIABLEDEB')),11,'M'))));
      SetControlText('PHC_DATEFIN',DateToStr(FinDeMois(PlusDate(StrToDate(GetControlText('PHC_DATEDEBUT')),11,'M'))));
// f PT1-1
    end;
  end;
end;

procedure TOF_PGEVOLUTIONEFFECTIF.ControlPeriode(Sender: TObject);
begin
PgIFValidPeriode;
end;

procedure TOF_PGEVOLUTIONEFFECTIF.ChangeExercice(Sender: TObject);
var
QPeriode : TQuery;
begin
QPeriode:=OpenSql('SELECT MIN(PPU_DATEDEBUT),MAX(PPU_DATEFIN),'+
                  'MAX(PEX_ANNEEREFER) FROM PAIEENCOURS '+
  'LEFT JOIN EXERSOCIAL ON PPU_DATEDEBUT>=PEX_DATEDEBUT AND PPU_DATEFIN<=PEX_DATEFIN '+
  'WHERE PEX_EXERCICE="'+GetControlText('EDEXERSOC')+'"',TRUE);
if not QPeriode.eof then
  Begin
// d PT1-1
//  SetControlText('XX_VARIABLEDEB',DateToStr(QPeriode.Fields[0].AsDateTime));
//  SetControlText('XX_VARIABLEFIN',DateToStr(QPeriode.Fields[1].AsDateTime));
  SetControlText('PHC_DATEDEBUT',DateToStr(QPeriode.Fields[0].AsDateTime));
  SetControlText('PHC_DATEFIN',DateToStr(QPeriode.Fields[1].AsDateTime));
  Annee :=  QPeriode.Fields[2].asstring;
  End
else
  Begin
//  SetControlText('XX_VARIABLEDEB',DateToStr(idate1900));
//  SetControlText('XX_VARIABLEFIN',DateToStr(idate1900));
  SetControlText('PHC_DATEDEBUT',DateToStr(idate1900));
  SetControlText('PHC_DATEFIN',DateToStr(idate1900));
// f PT1-1
  Annee := '';
  End;
  Ferme(QPeriode);
end;

Procedure TOF_PGEVOLUTIONEFFECTIF.ConstruireTob ();
var
  Etablu, Categlu               : string;
begin
  Etablu := TobEff.getValue('PHC_ETABLISSEMENT');
  Categlu := TobEff.getValue('CATEGORIE');
  if (Etablu <> Etab) or  (Categ = '') then
  begin
    T := Tob.Create('Les effectif par catégorie',TobEtat,-1);
    Etab := TobEff.getValue('PHC_ETABLISSEMENT');
    Categ := 'xx';
    T.AddChampSupValeur('PHC_ETABLISSEMENT',Etab, False);
// d PT1-1
    T.AddChampSupValeur('PHC_DATEDEBUT',GetControlText('PHC_DATEDEBUT'), False);
    T.AddChampSupValeur('PHC_DATEFIN',GetControlText('PHC_DATEFIN'), False);
// f PT1-1
    T.AddChampSupValeur('LIBETAB',TobEff.getValue('ET_LIBELLE'), False);
    T.AddChampSupValeur('Ouv01',0,FALSE);
    T.AddChampSupValeur('Ouv02',0,FALSE);
    T.AddChampSupValeur('Ouv03',0,FALSE);
    T.AddChampSupValeur('Ouv04',0,FALSE);
    T.AddChampSupValeur('Ouv05',0,FALSE);
    T.AddChampSupValeur('Ouv06',0,FALSE);
    T.AddChampSupValeur('Ouv07',0,FALSE);
    T.AddChampSupValeur('Ouv08',0,FALSE);
    T.AddChampSupValeur('Ouv09',0,FALSE);
    T.AddChampSupValeur('Ouv10',0,FALSE);
    T.AddChampSupValeur('Ouv11',0,FALSE);
    T.AddChampSupValeur('Ouv12',0,FALSE);
    T.AddChampSupValeur('Agm01',0,FALSE);
    T.AddChampSupValeur('Agm02',0,FALSE);
    T.AddChampSupValeur('Agm03',0,FALSE);
    T.AddChampSupValeur('Agm04',0,FALSE);
    T.AddChampSupValeur('Agm05',0,FALSE);
    T.AddChampSupValeur('Agm06',0,FALSE);
    T.AddChampSupValeur('Agm07',0,FALSE);
    T.AddChampSupValeur('Agm08',0,FALSE);
    T.AddChampSupValeur('Agm09',0,FALSE);
    T.AddChampSupValeur('Agm10',0,FALSE);
    T.AddChampSupValeur('Agm11',0,FALSE);
    T.AddChampSupValeur('Agm12',0,FALSE);
    T.AddChampSupValeur('Tec01',0,FALSE);
    T.AddChampSupValeur('Tec02',0,FALSE);
    T.AddChampSupValeur('Tec03',0,FALSE);
    T.AddChampSupValeur('Tec04',0,FALSE);
    T.AddChampSupValeur('Tec05',0,FALSE);
    T.AddChampSupValeur('Tec06',0,FALSE);
    T.AddChampSupValeur('Tec07',0,FALSE);
    T.AddChampSupValeur('Tec08',0,FALSE);
    T.AddChampSupValeur('Tec09',0,FALSE);
    T.AddChampSupValeur('Tec10',0,FALSE);
    T.AddChampSupValeur('Tec11',0,FALSE);
    T.AddChampSupValeur('Tec12',0,FALSE);
    T.AddChampSupValeur('Emp01',0,FALSE);
    T.AddChampSupValeur('Emp02',0,FALSE);
    T.AddChampSupValeur('Emp03',0,FALSE);
    T.AddChampSupValeur('Emp04',0,FALSE);
    T.AddChampSupValeur('Emp05',0,FALSE);
    T.AddChampSupValeur('Emp06',0,FALSE);
    T.AddChampSupValeur('Emp07',0,FALSE);
    T.AddChampSupValeur('Emp08',0,FALSE);
    T.AddChampSupValeur('Emp09',0,FALSE);
    T.AddChampSupValeur('Emp10',0,FALSE);
    T.AddChampSupValeur('Emp11',0,FALSE);
    T.AddChampSupValeur('Emp12',0,FALSE);
    T.AddChampSupValeur('Cad01',0,FALSE);
    T.AddChampSupValeur('Cad02',0,FALSE);
    T.AddChampSupValeur('Cad03',0,FALSE);
    T.AddChampSupValeur('Cad04',0,FALSE);
    T.AddChampSupValeur('Cad05',0,FALSE);
    T.AddChampSupValeur('Cad06',0,FALSE);
    T.AddChampSupValeur('Cad07',0,FALSE);
    T.AddChampSupValeur('Cad08',0,FALSE);
    T.AddChampSupValeur('Cad09',0,FALSE);
    T.AddChampSupValeur('Cad10',0,FALSE);
    T.AddChampSupValeur('Cad11',0,FALSE);
    T.AddChampSupValeur('Cad12',0,FALSE);
    T.AddChampSupValeur('Aut01',0,FALSE);
    T.AddChampSupValeur('Aut02',0,FALSE);
    T.AddChampSupValeur('Aut03',0,FALSE);
    T.AddChampSupValeur('Aut04',0,FALSE);
    T.AddChampSupValeur('Aut05',0,FALSE);
    T.AddChampSupValeur('Aut06',0,FALSE);
    T.AddChampSupValeur('Aut07',0,FALSE);
    T.AddChampSupValeur('Aut08',0,FALSE);
    T.AddChampSupValeur('Aut09',0,FALSE);
    T.AddChampSupValeur('Aut10',0,FALSE);
    T.AddChampSupValeur('Aut11',0,FALSE);
    T.AddChampSupValeur('Aut12',0,FALSE);
  end;
  //DEBUT PT2
  if (Categlu = '01') OR (Categlu = 'ZZZ') then
  begin
    case TobEff.GetValue('MOIS') of
    1 : T.PutValue ('Ouv01',T.GetValue('Ouv01') + TobEff.GetValue('EFF'));
    2 : T.PutValue ('Ouv02',T.GetValue('Ouv02') + TobEff.GetValue('EFF'));
    3 : T.PutValue ('Ouv03',T.GetValue('Ouv03') + TobEff.GetValue('EFF'));
    4 : T.PutValue ('Ouv04',T.GetValue('Ouv04') + TobEff.GetValue('EFF'));
    5 : T.PutValue ('Ouv05',T.GetValue('Ouv05') + TobEff.GetValue('EFF'));
    6 : T.PutValue ('Ouv06',T.GetValue('Ouv06') + TobEff.GetValue('EFF'));
    7 : T.PutValue ('Ouv07',T.GetValue('Ouv07') + TobEff.GetValue('EFF'));
    8 : T.PutValue ('Ouv08',T.GetValue('Ouv08') + TobEff.GetValue('EFF'));
    9 : T.PutValue ('Ouv09',T.GetValue('Ouv09') + TobEff.GetValue('EFF'));
   10 : T.PutValue ('Ouv10',T.GetValue('Ouv10') + TobEff.GetValue('EFF'));
   11 : T.PutValue ('Ouv11',T.GetValue('Ouv11') + TobEff.GetValue('EFF'));
   12 : T.PutValue ('Ouv12',T.GetValue('Ouv12') + TobEff.GetValue('EFF'));
   end;
  end
  else
    if (Categlu = '04') then
    begin
      case TobEff.GetValue('MOIS') of
      1 : T.PutValue ('Agm01',TobEff.GetValue('EFF'));
      2 : T.PutValue ('Agm02',TobEff.GetValue('EFF'));
      3 : T.PutValue ('Agm03',TobEff.GetValue('EFF'));
      4 : T.PutValue ('Agm04',TobEff.GetValue('EFF'));
      5 : T.PutValue ('Agm05',TobEff.GetValue('EFF'));
      6 : T.PutValue ('Agm06',TobEff.GetValue('EFF'));
      7 : T.PutValue ('Agm07',TobEff.GetValue('EFF'));
      8 : T.PutValue ('Agm08',TobEff.GetValue('EFF'));
      9 : T.PutValue ('Agm09',TobEff.GetValue('EFF'));
     10 : T.PutValue ('Agm10',TobEff.GetValue('EFF'));
     11 : T.PutValue ('Agm11',TobEff.GetValue('EFF'));
     12 : T.PutValue ('Agm12',TobEff.GetValue('EFF'));
      end;
    end
    else
      if (Categlu = '03') then
      begin
        case TobEff.GetValue('MOIS') of
        1 : T.PutValue ('Tec01',TobEff.GetValue('EFF'));
        2 : T.PutValue ('Tec02',TobEff.GetValue('EFF'));
        3 : T.PutValue ('Tec03',TobEff.GetValue('EFF'));
        4 : T.PutValue ('Tec04',TobEff.GetValue('EFF'));
        5 : T.PutValue ('Tec05',TobEff.GetValue('EFF'));
        6 : T.PutValue ('Tec06',TobEff.GetValue('EFF'));
        7 : T.PutValue ('Tec07',TobEff.GetValue('EFF'));
        8 : T.PutValue ('Tec08',TobEff.GetValue('EFF'));
        9 : T.PutValue ('Tec09',TobEff.GetValue('EFF'));
       10 : T.PutValue ('Tec10',TobEff.GetValue('EFF'));
       11 : T.PutValue ('Tec11',TobEff.GetValue('EFF'));
       12 : T.PutValue ('Tec12',TobEff.GetValue('EFF'));
        end;
      end
      else
        if (Categlu = '02') then
        begin
          case TobEff.GetValue('MOIS') of
          1 : T.PutValue ('Emp01',TobEff.GetValue('EFF'));
          2 : T.PutValue ('Emp02',TobEff.GetValue('EFF'));
          3 : T.PutValue ('Emp03',TobEff.GetValue('EFF'));
          4 : T.PutValue ('Emp04',TobEff.GetValue('EFF'));
          5 : T.PutValue ('Emp05',TobEff.GetValue('EFF'));
          6 : T.PutValue ('Emp06',TobEff.GetValue('EFF'));
          7 : T.PutValue ('Emp07',TobEff.GetValue('EFF'));
          8 : T.PutValue ('Emp08',TobEff.GetValue('EFF'));
          9 : T.PutValue ('Emp09',TobEff.GetValue('EFF'));
         10 : T.PutValue ('Emp10',TobEff.GetValue('EFF'));
         11 : T.PutValue ('Emp11',TobEff.GetValue('EFF'));
         12 : T.PutValue ('Emp12',TobEff.GetValue('EFF'));
          end;
        end
        else
          if (Categlu = '12') or (Categlu = '13') then
          begin
            case TobEff.GetValue('MOIS') of
            1 : T.PutValue ('Cad01',T.GetValue('Cad01') + TobEff.GetValue('EFF'));
            2 : T.PutValue ('Cad02',T.GetValue('Cad02') + TobEff.GetValue('EFF'));
            3 : T.PutValue ('Cad03',T.GetValue('Cad03') + TobEff.GetValue('EFF'));
            4 : T.PutValue ('Cad04',T.GetValue('Cad04') + TobEff.GetValue('EFF'));
            5 : T.PutValue ('Cad05',T.GetValue('Cad05') + TobEff.GetValue('EFF'));
            6 : T.PutValue ('Cad06',T.GetValue('Cad06') + TobEff.GetValue('EFF'));
            7 : T.PutValue ('Cad07',T.GetValue('Cad07') + TobEff.GetValue('EFF'));
            8 : T.PutValue ('Cad08',T.GetValue('Cad08') + TobEff.GetValue('EFF'));
            9 : T.PutValue ('Cad09',T.GetValue('Cad09') + TobEff.GetValue('EFF'));
           10 : T.PutValue ('Cad10',T.GetValue('Cad10') + TobEff.GetValue('EFF'));
           11 : T.PutValue ('Cad11',T.GetValue('Cad11') + TobEff.GetValue('EFF'));
           12 : T.PutValue ('Cad12',T.GetValue('Cad12') + TobEff.GetValue('EFF'));
            end;
          end
          else
            if (Categlu > '04') and (Categlu <> '12') and (Categlu <> '13') then
            begin
              case TobEff.GetValue('MOIS') of
              1 : T.PutValue ('Aut01',T.GetValue('Aut01') + TobEff.GetValue('EFF'));
              2 : T.PutValue ('Aut02',T.GetValue('Aut02') + TobEff.GetValue('EFF'));
              3 : T.PutValue ('Aut03',T.GetValue('Aut03') + TobEff.GetValue('EFF'));
              4 : T.PutValue ('Aut04',T.GetValue('Aut04') + TobEff.GetValue('EFF'));
              5 : T.PutValue ('Aut05',T.GetValue('Aut05') + TobEff.GetValue('EFF'));
              6 : T.PutValue ('Aut06',T.GetValue('Aut06') + TobEff.GetValue('EFF'));
              7 : T.PutValue ('Aut07',T.GetValue('Aut07') + TobEff.GetValue('EFF'));
              8 : T.PutValue ('Aut08',T.GetValue('Aut08') + TobEff.GetValue('EFF'));
              9 : T.PutValue ('Aut09',T.GetValue('Aut09') + TobEff.GetValue('EFF'));
             10 : T.PutValue ('Aut10',T.GetValue('Aut10') + TobEff.GetValue('EFF'));
             11 : T.PutValue ('Aut11',T.GetValue('Aut11') + TobEff.GetValue('EFF'));
             12 : T.PutValue ('Aut12',T.GetValue('Aut12') + TobEff.GetValue('EFF'));
              end;
            end;
            //FIN PT2
  end;

Initialization
  registerclasses ( [ TOF_PGEVOLUTIONEFFECTIF ] ) ;
end.
