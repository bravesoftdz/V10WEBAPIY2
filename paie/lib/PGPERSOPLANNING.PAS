unit PGPersoPlanning;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, HColor, HTB97, StdCtrls, Grids, Hctrls, Spin, Mask{$IFNDEF DBXPRESS} ,dbTables {$ELSE} ,uDbxDataSet {$ENDIF};

type
  TPersoPlanning = class(TForm)
    Pages: TPageControl;
    TBCRITERE: TTabSheet;
    TBETIQUETTE: TTabSheet;
    Dock971: TDock97;
    PBouton: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    PPL_PGETQCHAMP: THValComboBox;
    TPPL_PGETQCHAMP: THLabel;
    TPPL_PGETQTEXTALIGN: THLabel;
    PPL_PGETQTEXTALIGN: THValComboBox;
    PPL_PGETQFORME: THValComboBox;
    TPPL_PGETQFORME: THLabel;
    GBCOLORFOND: TGroupBox;
    PPL_PGCOLLUNDI: TPaletteButton97;
    TPPL_PGCOLLUNDI: TLabel;
    TPPL_COLMARDI: TLabel;
    PPL_COLMARDI: TPaletteButton97;
    PPL_PGCOLMERCREDI: TPaletteButton97;
    TPPL_PGCOLMERCREDI: TLabel;
    PPL_PGCOLJEUDI: TPaletteButton97;
    TPPL_PGCOLJEUDI: TLabel;
    PPL_PGCOLVENDREDI: TPaletteButton97;
    TPPL_PGCOLVENDREDI: TLabel;
    PPL_PGCOLSAMEDI: TPaletteButton97;
    TPPL_PGCOLSAMEDI: TLabel;
    PPL_PGCOLDIMANCHE: TPaletteButton97;
    TPPL_PGCOLDIMANCHE: TLabel;
    PPL_PGCOLFERIE: TPaletteButton97;
    TPPL_PGCOLFERIE: TLabel;
    PPL_COLBKGROUND: TPaletteButton97;
    TPPL_COLBKGROUND: TLabel;
    TPPL_COLSELECTION: TLabel;
    PPL_COLSELECTION: TPaletteButton97;
    GBPARAM: TGroupBox;
    TPPL_PGINTERVAL: TLabel;
    PPL_PGINTERVAL: TComboBox;
    PPL_PGDATEFORMAT: THCritMaskEdit;
    PPL_PGSIZECOL: TSpinEdit;
    TPPL_PGSIZECOL: THLabel;
    PPL_PGSIZEROW: TSpinEdit;
    TPPL_PGSIZEROW: THLabel;
    PPL_PGETQBKGCOLOR: TPaletteButton97;
    procedure BFermeClick(Sender: TObject);
  private
    { Déclarations privées }
    procedure LoadParamdefaut;
  public
    { Déclarations publiques }
    StTypUtilisateur,StSalarie,StPgTypePlanning : string;
  end;

  procedure PGPersoPlanningAbsence(TypUtilisateur,Salarie,PgTypePlanning : string) ;


var
  PersoPlanning : TPersoPlanning;

implementation

{$R *.DFM}

procedure PGPersoPlanningAbsence(TypUtilisateur,Salarie,PgTypePlanning : string) ;
Begin
  PersoPlanning:=nil;
  PersoPlanning := TPersoPlanning.Create(Application);
  try
    PersoPlanning.StTypUtilisateur  := TypUtilisateur;
    PersoPlanning.StSalarie         := Salarie;
    PersoPlanning.StPgTypePlanning  := PgTypePlanning;
    PersoPlanning.ShowModal ;
  finally
    if PersoPlanning<>nil then PersoPlanning.Free;
  end;

End;


procedure TPersoPlanning.BFermeClick(Sender: TObject);
begin
PersoPlanning.Close;
end;

procedure TPersoPlanning.LoadParamdefaut;
Var
Q : TQuery;
i : Integer;
begin
i := 0 ;
if StPgTypePlanning='ECONGES' then
  Begin
  PPL_PGEtqChamp.value := 'RES'; //Etat de la validation responsable
  { Valeur par défaut des états => grille }
  Q := OpenSql('SELECT * FROM COMMUN WHERE CO_TYPE="PAE" AND CO_CODE<>"REP"',True);
  While not Q.Eof Do
    Begin
    Inc(i);

    {Grille.CellValues[0,i] := Q.FindField('CO_LIBELLE').AsString;
    if Q.FindField('CO_CODE').AsString = 'VAL' then
    Grille.CellValues[1,i] := Q.FindField('CO_LIBELLE').AsString;
    Grille.CellValues[2,i] := Q.FindField('CO_LIBELLE').AsString;
    Grille.CellValues[3,i] := Q.FindField('CO_LIBELLE').AsString;
    Grille.CellValues[4,i] := Q.FindField('CO_LIBELLE').AsString;

    T.PutValue('BACKGROUNDCOLOR',BackGroundColor);
           if (T.getValue('CO_CODE')='VAL') then
          BackGroundColor:='10944422'
       else
          if (T.getValue('CO_CODE')='ATT') then
             BackGroundColor:='10210815'//'16053248'
          else
             BackGroundColor:='8421631';
       End


    T.PutValue('FONTCOLOR',FontColor);
    T.PutValue('FONTNAME',FontName);
    T.PutValue('FONTSTYLE',FontStyle);
    T.PutValue('FONTSIZE',FontSize);
      }

    Q.Next;
    End;
  Ferme(Q);

  {

  FontColor:='ClBlack';
  FontStyle :='G';
  FontName:='Times New Roman';
  FontSize:=10;
  ColorBackground:='clWhite';
  H.Interval       := piJour; // ;piDemiJour; //
  H.CumulInterval  := pciSemaine ;
  H.IntervalDebut := Planning.DateDebAbs;
  H.IntervalFin   := Planning.DateFinAbs;
  H.MultiLine      := False ;
  H.GestionJoursFeriesActive := True ;
  H.ActiveSaturday := True ;
  H.ActiveSunday   := True ;
  H.ColorJoursFeries  := StringToColor('12189695');
  H.ColorOfSaturday := StringToColor('12189695');
  H.ColorOfSunday   := StringToColor('12189695');
  H.ColorBackground :=StringToColor(ColorBackground);
  H.ColorSelection:=clTeal;//StringToColor(FontColor);
  H.ActiveLigneDate := True ;
  H.ActiveLigneGroupeDate := True ;
  H.JourneeDebut := StrToTime('07:00') ;
  H.JourneeFin   := StrToTime('19:00') ;
  H.DateFormat   := 'dd mm' ;  //hh:mm
  H.DebutAM      := StrToTime('08:00') ;
  H.FinAM        := StrToTime('12:00') ;
  H.DebutAP      := StrToTime('14:00') ;
  H.FinAP        := StrToTime('18:00') ;     }

  end;

end;

end.
