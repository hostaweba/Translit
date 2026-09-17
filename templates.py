# templates.py
"""
Premium HTML Templates for the Translit Application
"""

# Replaced primitive underscores with a clean, print-safe horizontal CSS rule.
U_LINE = "<hr style='border: none; border-bottom: 2px solid #374151; margin: 15px 0;'>"

TPL_QUESTION_PAPER = f"""
<div style="font-family: Arial, sans-serif;">
    <div style="text-align: center; margin-bottom: 15px;">
        <h1 style="margin: 0; font-size: 24pt; color: #111827;">[विद्यालय / संस्था का नाम]</h1>
        <h3 style="margin: 5px 0 10px 0; font-size: 16pt; color: #4b5563; font-weight: normal;">वार्षिक परीक्षा २०२४-२५</h3>
    </div>
    
    <table width="100%" cellpadding="6" cellspacing="0" style="border-top: 2px solid #000; border-bottom: 2px solid #000; margin-bottom: 15px;">
        <tr>
            <td align="left" width="50%" style="font-size: 12pt;"><b>कक्षा:</b> ........................</td>
            <td align="right" width="50%" style="font-size: 12pt;"><b>विषय:</b> ........................</td>
        </tr>
        <tr>
            <td align="left" width="50%" style="font-size: 12pt;"><b>समय:</b> ३ घंटे</td>
            <td align="right" width="50%" style="font-size: 12pt;"><b>पूर्णांक:</b> १००</td>
        </tr>
    </table>
    
    <div style="background-color: #f3f4f6; padding: 10px; border: 1px solid #d1d5db; margin-bottom: 20px;">
        <b style="color: #111827;">निर्देश:</b> सभी प्रश्न अनिवार्य हैं। प्रत्येक प्रश्न के निर्धारित अंक उसके सामने कोष्ठक में दिए गए हैं।
    </div>
    
    <ol style="font-size: 12pt; line-height: 1.8;">
        <li>यहाँ अपना पहला प्रश्न लिखें। (५ अंक)</li>
        <li>यहाँ अपना दूसरा प्रश्न विस्तार से लिखें। (५ अंक)</li>
        <li>निम्नलिखित में से किन्हीं दो के उत्तर दें: (१० अंक)</li>
    </ol>
</div>
"""

TPL_APPLICATION = """
<div style="font-size: 12pt; line-height: 1.6;">
    <p><b>सेवा में,</b></p>
    <p style="margin-left: 30px;">
        श्रीमान प्रधानाचार्य महोदय,<br>
        <b>[विद्यालय / संस्था का नाम]</b>,<br>
        [शहर, राज्य का नाम]
    </p>
    
    <p style="text-align: center; font-weight: bold; text-decoration: underline; font-size: 14pt; margin: 25px 0;">
        विषय: [अवकाश / अन्य हेतु आवेदन पत्र]
    </p>
    
    <p><b>महोदय,</b></p>
    <p style="text-indent: 50px; text-align: justify;">
        सविनय निवेदन है कि [यहाँ अपना कारण लिखें - उदाहरण: मुझे कल रात से तेज बुखार आ रहा है, जिसके कारण मैं विद्यालय उपस्थित होने में असमर्थ हूँ]। अतः आपसे विनम्र अनुरोध है कि मुझे [प्रारंभ तिथि] से [समाप्ति तिथि] तक का अवकाश प्रदान करने की कृपा करें।
    </p>
    <p style="text-indent: 50px;">
        इसके लिए मैं सदैव आपका आभारी रहूँगा / रहूँगी।
    </p>
    
    <table width="100%" border="0" cellpadding="5" style="margin-top: 50px;">
        <tr>
            <td align="left" valign="bottom">
                <b>दिनांक:</b> ....../....../२०....<br>
                <b>स्थान:</b> ........................
            </td>
            <td align="right" valign="bottom">
                <b>आपका आज्ञाकारी शिष्य / शिष्या</b><br><br>
                <b>नाम:</b> ........................<br>
                <b>कक्षा:</b> ........................<br>
                <b>अनुक्रमांक:</b> ........................
            </td>
        </tr>
    </table>
</div>
"""

TPL_DECLARATION = """
<div style="font-size: 12pt; line-height: 1.6;">
    <h2 style="text-align: center; font-size: 20pt; text-decoration: underline; margin-bottom: 30px; letter-spacing: 1px;">स्व-घोषणा पत्र (Declaration)</h2>
    
    <p style="text-align: justify;">
        मैं, <b>[आपका नाम]</b>, पुत्र/पुत्री/पत्नी श्री <b>[पिता/पति का नाम]</b>, आयु <b>[उम्र]</b> वर्ष, निवासी <b>[पूरा पता]</b>, एतद्द्वारा सत्यनिष्ठा से यह घोषणा करता/करती हूँ कि:
    </p>
    
    <ol style="margin-top: 20px; margin-bottom: 20px;">
        <li style="margin-bottom: 12px; text-align: justify;">मेरे द्वारा उपरोक्त आवेदन / प्रपत्र में दी गई सभी जानकारी मेरी व्यक्तिगत जानकारी और विश्वास के अनुसार पूर्णतः सत्य और सही है।</li>
        <li style="margin-bottom: 12px; text-align: justify;">मैंने कोई भी तथ्य या आवश्यक जानकारी नहीं छुपाई है तथा कोई भी झूठा साक्ष्य प्रस्तुत नहीं किया है।</li>
        <li style="margin-bottom: 12px; text-align: justify;">यदि भविष्य में किसी भी स्तर पर मेरे द्वारा दी गई कोई भी जानकारी असत्य या भ्रामक पाई जाती है, तो संस्था / विभाग को मेरे विरुद्ध वैधानिक कार्यवाही करने का पूर्ण अधिकार होगा और इसके लिए मैं स्वयं जिम्मेदार रहूँगा/रहूँगी।</li>
    </ol>
    
    <table width="100%" border="0" cellpadding="5" style="margin-top: 60px;">
        <tr>
            <td align="left" valign="bottom">
                <b>स्थान:</b> ........................<br>
                <b>दिनांक:</b> ....../....../२०....
            </td>
            <td align="right" valign="bottom">
                ..................................................<br>
                <b>(हस्ताक्षर घोषणाकर्ता)</b><br><br>
                <b>नाम:</b> ........................................<br>
                <b>मोबाइल:</b> ........................................
            </td>
        </tr>
    </table>
</div>
"""