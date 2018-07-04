Unit UTofAfTarifart_Mul ;

Interface

Uses
{$IFDEF EAGLCLIENT}
      Maineagl,
{$ELSE}
      Fe_Main,
{$ENDIF}
    forms,stdctrls,Classes,  HCtrls,
    AfUtilArticle, HEnt1,
    UtofTarifArt_mul,UtilGc  ;

Type
  Tof_AfTarifArt_Mul = Class (Tof_TarifArt_Mul)
   public
    { Déclarations publiques }

   private
    { Déclarations privées }

    procedure Onargument(stArgument : String ) ; override;
    end;



Procedure AFLanceFiche_Mul_TarifArt(Range,Argument:string);

implementation

procedure Tof_AfTarifArt_Mul.OnArgument(stArgument : String ) ;
var     TLab : THLabel;
    ComboTypeArticle : THValComboBox;
begin
InHerited;
                //article
if ((Ecran.FindComponent ('GA_LIBREART1'))<>nil) then
    begin
    GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'GA_LIBREART', 10, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GA_VALLIBRE', 3, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GA_DATELIBRE', 3, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GA_CHARLIBRE', 3, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'BOOL', 'GA_BOOLLIBRE', 3, '_');
    end;

TLab:=THLabel (Ecran.FindComponent ('TGA_FAMILLENIV1'));
if tlab <>NIL then SetCOntrolText('TGA_FAMILLENIV1', RechDom('GCLIBFAMILLE','LF1',false));
TLab:=THLabel (Ecran.FindComponent ('TGA_FAMILLENIV2'));
if tlab <>NIL then SetCOntrolText('TGA_FAMILLENIV2', RechDom('GCLIBFAMILLE','LF2',false));
TLab:=THLabel (Ecran.FindComponent ('TGA_FAMILLENIV3'));
if tlab <>NIL then SetCOntrolText('TGA_FAMILLENIV3', RechDom('GCLIBFAMILLE','LF3',false));
  //mcd 05/03/02
ComboTypeArticle:=THValComboBox(GetControl('GA_TYPEARTICLE'));
ComboTypeArticle.plus:=PlusTypeArticle;

end;
Procedure AFLanceFiche_Mul_TarifArt(Range,Argument:string);
begin
AGLLanceFiche ('AFF','AFTARIFART_MUL',range,'',Argument);
end;


Initialization
  registerclasses ( [ TOF_AFTARIFART_MUL ] ) ; 
end.
