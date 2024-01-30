from fpdf import FPDF

def from_mmc(mmc, pn):
    if mmc[0:3] == 'ADR': #S10
        encoded = {'pos1': 0}




if __name__ == "__main__":
    header = """WHITE DRIVE MOTORS & STEERING          PAGE 1
    LOGISTYCZNA 1, BIELANY WROCLAWSKIE, 55-040 KOBIERZYCE, POLAND
    ENGINEERING SPECIFICATION SHEET"""
    t = """PRODUCT LINE : SCU             REVISION   : E
    PRODUCT NO.  : 200-0002-002    MODEL CODE : ADRA21A51ADF00122400AACAA2AAB1FB
    INSTALLATION : A-1372-001                   0....0....1....1....2....2....3....3
    --------------------------------------------0----5----0----5----0----5----0----5
    DOCUMENT BASED ON ANSI Y14.5M-1982 ALL DIMENSIONS ARE IN MILLIMETERS[INCHES].
    DESCRIPTION:
    ADR  PRODUCT - SERIES 10
    A    UNIT TYPE - STANDARD
    2    FLOW RATING - 3.8-30 L/MIN [1.00-8.00 GAL/MIN] CLOSED CENTER, LOAD SENSING
    1    INLET PRESSURE RATING - 276 BAR [4000 LBF/IN2]
    A    RETURN PRESSURE RATING - 21 BAR [305 LBF/IN2] MAXIMUM
    51   DISPLACEMENT - 159 CM3/R [9.73 IN3/R]
    A    FLOW AMPLIFICATION - NONE
    D    NEUTRAL CIRCUIT - LOAD SENSING, STATIC SIGNAL
    F    LOAD CIRCUIT - MODIFIED LOAD REACTION
    00   SPECIAL SPOOL\SLEEVE MODIFICATION - NONE
    12   VALVE OPTIONS - ANTI-CAVITATION PRELOAD VALVE, INLET CHECK VALVE, LOAD
         SENSING RELIEF VALVE, MANUAL STEERING CHECK VALVE
    24   INLET OR LOAD SENSE RELIEF VALVE SETTING - 165 +7/-0 BAR [2390 +100/-0
         LBF/IN2]
    00   CYLINDER RELIEF VALVE SETTING - NONE
    A    P,T,L AND R PORT SIZE - 4X .750-16 UNF-2B SAE O-RING PORT
    AC   ADDITIONAL PORTS - .4375-20 UNF-2B LOAD SENSING SAE O-RING PORT ON PORT
         FACE
    A    MOUNTING THREADS - 2X M12 X 1.75-6H X 24.1 [.95] MIN DEPTH PORT FACE 4X M10
         X 1.5-6H MOUNTING FACE
    A    MECHANICAL INTERFACE - INTERNAL INVOLUTE SPLINE, 12 TOOTH 16/32 DP 30
         DEGREE PA
    2    INPUT TORQUE - MEDIUM
    A    FLUID TYPE - REFERENCE EATON TECHNICAL BULLETIN 3-401
    AB   SPECIAL FEATURES - WHITE PAINT DOT ON HOUSING
    1    PAINTS & PACKAGING - BLACK PAINT
    F    IDENTIFICATION - CASE NEW HOLLAND PART NUMBER ON NAMEPLATE
    B    DESIGN CODE - 002
    Customer Name: CASE NEW HOLLAND
    Customer Number: 37787A1
    NOTES:
    SERVICE INFORMATION:
    SEAL KIT NUMBER: 9900287-000
    NOTICE: WHITE DRIVE MOTORS & STEERING PROPRIETARY INFORMATION"""

    pdf = FPDF()
    pdf.add_page()
    pdf.add_font('Consolas', 'B', r'C:\Windows\Fonts\consola.ttf', uni=True)
    pdf.set_font('Consolas', 'B', size=11)

    numbers = ['A', 'D']

    column_t = []

    for i in header.splitlines():
        pdf.cell(0, 4, txt=i, ln=1, align='C')

    # for i in t.splitlines():
    #     pdf.cell(0, 4, txt=i, ln=1, align='L')

    for i in numbers:
        pdf.cell(20, 4, txt=i, ln=1, align='L')

    for i in column_t:
        pdf.cell(0, 4, txt=i, ln=1, align='L')

    pdf.output(r"C:\Users\u331609\Desktop\mygfg.pdf")