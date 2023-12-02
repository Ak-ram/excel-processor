exports.mergeDestinations = [
    "وحدة الخدمات", "مفتش الدخلية", "ادارة شئون الخدمة", "م.المدير للوحدات",
    "م.المدير للشؤن المالية", "م.المدير للافراد والتدريب", "1قسم العلاقات",
    "1قسم الانضباط", "1قسم الاسلحة", "1قسم المعلومات والتوثيق", "1قسم الانشاءات",
    "م.المدير للامن العام", "1قسم التخطيط والمتابعة", "1قسم التحقيقات", "1قسم الرخص",
    "نائب م.الامن", "1قسم الرقابة الجنائية", "1قسم حقوق الانسان",
  ];

exports.combinations = {
    'بوفيه': 'بوفيه_مستشفى',
    'مستشفى': 'بوفيه_مستشفى',
    'مباحث الادارة': 'مباحث_سياسين',
    'السياسين': 'مباحث_سياسين',
  };


exports.style_data = {
  fontSize: 18,
  bold: true,
  horizontalAlignment: "center",
  verticalAlignment: "center",
  border: true,
  shrinkToFit: true,
};

exports.style_header = {
  fontSize: 20,
  bold: true,
  horizontalAlignment: "center",
  verticalAlignment: "center",
  border: true,
  fill: {
    type: "solid",
    color: "D9D9D9",
  },
  shrinkToFit: true,
};
