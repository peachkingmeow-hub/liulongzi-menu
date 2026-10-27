/**
 * ===== 菜单数据文件 menu-data.js =====
 *
 * 更新方式：
 *   方式1（推荐）：运行 python convert-excel.py <Excel文件路径>
 *       例：python convert-excel.py 商品+套餐导出文件.xlsx
 *       会自动读取 Excel 并重新生成此文件
 *
 *   方式2（手动）：直接编辑此文件中的 MENU 数组
 *
 * 数据格式说明：
 *   MENU = [
 *     {
 *       id:    "cat1",              // 分类ID（固定，勿改）
 *       name:  "招牌牛肉粉面",        // 分类名称
 *       items: [
 *         {
 *           id:    "p1101",         // 商品ID（格式 p + 分类号 + 序号）
 *           name:  "招牌红烧牛肉粉",  // 商品名称
 *           price: 25.0,             // 基础价格（元）
 *           desc:  "描述文字",       // 商品描述
 *           tags:  ["微辣"],         // 标签数组
 *           img:   "https://...",   // 商品图片URL
 *           opts:  [                 // 选项组（可选）
 *             {
 *               g:    "粉面主食",    // 选项组名称
 *               req:  true,          // 是否必选
 *               opts: ["圆粉","扁粉"] // 可选项
 *             }
 *           ],
 *           specs: []                // 规格信息（可选）
 *         }
 *       ]
 *     }
 *   ]
 *
 * 注意：以下分类ID与餐盒逻辑绑定，修改后需同步调整 index.html：
 *   cat1  -> 招牌牛肉粉（干湿分离盒 2元）
 *   cat2  -> 招牌牛肉面（品牌餐盒 1元）
 *   cat5  -> 鲜鱼汤-现炒码子粉（码子粉干湿分离 2.5元）
 *   cat6  -> 汉川特色掺汤（品牌餐盒 1元）
 *   cat7  -> 海鲜粥（品牌餐盒 1元）
 *   cat8  -> 新品专区（品牌餐盒 1元）
 *   cat9  -> 鲜货现炒码子（品牌餐盒 1元）
 *   cat10 -> 特色炖锅（炖锅打包盒 2元）
 *   cat11 -> 鲜货炭火烧烤（打包带走 免费）
 *   cat12 -> 时令鲜汤（品牌餐盒 1元）
 *   cat13 -> 卤菜凉菜（码子方盒 0.5元）
 *   cat14 -> 炒粉炒面炒饭炒豆丝（品牌餐盒 1元）
 *   cat3/cat4/cat15/cat16（小吃、饮品、主食、酒水）-> 无默认餐盒
 */

const MENU = [
  {
    "id": "cat1",
    "name": "招牌牛肉粉",
    "items": [
      {
        "id": "p101",
        "name": "炸酱粉",
        "price": 12.0,
        "desc": "爽滑米粉裹满醇香肉酱，咸香浓郁、越拌越香。",
        "tags": [
          "做法原因不可续"
        ],
        "img": "https://s.koubei.com/4pSd8TZsXrnK5NmaXpAZQKen43MSB5jd.jpg",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "圆粉",
              "宽粉",
              "手工大宽粉",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "碱面",
              "豆丝",
              "湖南宽粉",
              "扁粉"
            ]
          },
          {
            "g": "加码小吃",
            "req": false,
            "opts": [
              "加-卤蛋/枚",
              "加-海带结2个",
              "加-卤鸡爪/只"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】品牌餐盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p102",
        "name": "招牌秘制牛肉粉",
        "price": 25.0,
        "desc": "文火慢卤切片，汤底醇厚，微辣鲜香",
        "tags": [
          "招牌必点"
        ],
        "img": "https://s.koubei.com/s8efjaBMyH2eKB3WicSYQ43CkPJPnRrA.jpg",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "圆粉",
              "宽粉",
              "手工大宽粉",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "碱面",
              "豆丝",
              "湖南宽粉",
              "扁粉"
            ]
          },
          {
            "g": "粉面加码",
            "req": false,
            "opts": [
              "加码-秘制牛肉"
            ]
          },
          {
            "g": "加码小吃",
            "req": false,
            "opts": [
              "加-卤蛋/枚",
              "加-海带结2个",
              "加-卤鸡爪/只"
            ]
          },
          {
            "g": "温馨提示",
            "req": false,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】干湿分离打包盒",
              "【打包】品牌餐盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p103",
        "name": "红烧牛肉粉",
        "price": 25.0,
        "desc": "微辣，慢炖坨状牛腩，软糯Q弹，老少皆宜",
        "tags": [],
        "img": "https://s.koubei.com/BjHyGWcXMk3QCii2TYy7zRa78iHAahQf.jpg",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "圆粉",
              "宽粉",
              "手工大宽粉",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "碱面",
              "豆丝",
              "湖南宽粉",
              "扁粉"
            ]
          },
          {
            "g": "粉面加码",
            "req": false,
            "opts": [
              "加码-红烧牛肉"
            ]
          },
          {
            "g": "加码小吃",
            "req": false,
            "opts": [
              "加-卤蛋/枚",
              "加-海带结2个",
              "加-卤鸡爪/只"
            ]
          },
          {
            "g": "温馨提示",
            "req": false,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】干湿分离打包盒",
              "【打包】品牌餐盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p104",
        "name": "特色牛三鲜粉",
        "price": 25.0,
        "desc": "脆爽与软糯交织，中辣过瘾",
        "tags": [],
        "img": "https://s.koubei.com/Qr7kQWJ4Tp262e3RnXTEFx5CR8GhBfAf.jpg",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "圆粉",
              "宽粉",
              "手工大宽粉",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "碱面",
              "豆丝",
              "湖南宽粉",
              "扁粉"
            ]
          },
          {
            "g": "粉面加码",
            "req": false,
            "opts": [
              "加码-牛三鲜"
            ]
          },
          {
            "g": "加码小吃",
            "req": false,
            "opts": [
              "加-卤蛋/枚",
              "加-海带结2个",
              "加-卤鸡爪/只"
            ]
          },
          {
            "g": "温馨提示",
            "req": false,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】干湿分离打包盒",
              "【打包】品牌餐盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p105",
        "name": "津市麻辣牛肉粉",
        "price": 28.0,
        "desc": "湖南特色风味，干香麻辣，嚼劲十足！",
        "tags": [
          "麻辣"
        ],
        "img": "https://s.koubei.com/zP6YkKPfwwtrykySKzs3ewGhEjmC7YQn.jpg",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "圆粉",
              "宽粉",
              "手工大宽粉",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "碱面",
              "豆丝",
              "湖南宽粉",
              "扁粉"
            ]
          },
          {
            "g": "加码小吃",
            "req": false,
            "opts": [
              "加-卤蛋/枚",
              "加-海带结2个",
              "加-卤鸡爪/只"
            ]
          },
          {
            "g": "温馨提示",
            "req": false,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】干湿分离打包盒",
              "【打包】品牌餐盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p106",
        "name": "鲜切原味牛肉粉",
        "price": 25.0,
        "desc": "大片牛肉卤香浓郁，配牛骨原汤，清淡鲜美，可选不辣或微辣",
        "tags": [
          "清汤"
        ],
        "img": "https://s.koubei.com/n65eksFbPccETkaBBJK5jXhPcZcmsjZS.jpg",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "圆粉",
              "宽粉",
              "手工大宽粉",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "碱面",
              "豆丝",
              "湖南宽粉",
              "扁粉"
            ]
          },
          {
            "g": "粉面加码",
            "req": false,
            "opts": [
              "加码-鲜切大片原味牛肉"
            ]
          },
          {
            "g": "加码小吃",
            "req": false,
            "opts": [
              "加-卤蛋/枚",
              "加-海带结2个",
              "加-卤鸡爪/只"
            ]
          },
          {
            "g": "辣度",
            "req": true,
            "opts": [
              "清汤不辣",
              "微辣"
            ]
          },
          {
            "g": "温馨提示",
            "req": false,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】品牌餐盒",
              "【打包】干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p107",
        "name": "热干系列",
        "price": 8.0,
        "desc": "默认微辣",
        "tags": [
          "做法原因不可续面"
        ],
        "img": "https://s.koubei.com/K8rS482afKkc2Cf6Kh4hdRSYCTzdTWct.jpg",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "热干面",
              "热干-宽粉",
              "热干-圆粉",
              "湖南宽粉"
            ]
          },
          {
            "g": "加码小吃",
            "req": false,
            "opts": [
              "加-卤蛋/枚",
              "加-海带结2个",
              "加-卤鸡爪/只"
            ]
          },
          {
            "g": "温馨提示",
            "req": false,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】品牌餐盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "原味热干粉",
            "price": 8.0
          },
          {
            "name": "原味热干面",
            "price": 8.0
          },
          {
            "name": "秘制牛肉热干粉",
            "price": 25.0
          },
          {
            "name": "秘制牛肉热干面",
            "price": 25.0
          }
        ]
      },
      {
        "id": "p108",
        "name": "清汤小伢粉",
        "price": 5.0,
        "desc": "消费满20元且带1.1米以下宝宝免费领取",
        "tags": [],
        "img": "https://s.koubei.com/RSdbYFbEH5G7ApMEFEX8Je3FtbGMAG6x.png",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "圆粉",
              "宽粉",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "碱面",
              "湖南宽粉",
              "扁粉"
            ]
          },
          {
            "g": "辣度",
            "req": true,
            "opts": [
              "清汤不辣"
            ]
          }
        ]
      },
      {
        "id": "p109",
        "name": "原汤粉/面",
        "price": 8.0,
        "desc": "原汤清鲜，粉面爽滑",
        "tags": [],
        "img": "https://s.koubei.com/QR7SZhNwtQkFY2TKpNMPbrhNA4RPfTBC.jpg",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "圆粉",
              "宽粉",
              "手工大宽粉",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "碱面",
              "豆丝",
              "湖南宽粉",
              "扁粉"
            ]
          },
          {
            "g": "加码小吃",
            "req": false,
            "opts": [
              "加-卤蛋/枚",
              "加-海带结2个",
              "加-卤鸡爪/只"
            ]
          },
          {
            "g": "辣度",
            "req": true,
            "opts": [
              "微辣",
              "清汤不辣"
            ]
          },
          {
            "g": "温馨提示",
            "req": false,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】品牌餐盒",
              "【打包】干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ]
      }
    ]
  },
  {
    "id": "cat2",
    "name": "特色小吃",
    "items": [
      {
        "id": "p201",
        "name": "牛肉汤卤蛋/枚",
        "price": 3.0,
        "desc": "牛肉原汤慢卤，咸香入味",
        "tags": [],
        "img": "https://s.koubei.com/eBem3bKsBhEPjrCasNeRRiGN5eBj6Hi6.jpg"
      },
      {
        "id": "p202",
        "name": "卤鸡爪/只",
        "price": 6.0,
        "desc": "软糯脱骨，卤香甜辣浓郁",
        "tags": [],
        "img": "https://s.koubei.com/jjMxhxp8GDjCJSiRiPJSWR3xWH2JSfMB.jpg"
      },
      {
        "id": "p203",
        "name": "卤海带结2个",
        "price": 3.0,
        "desc": "卤香甜辣入味，爽脆有嚼劲",
        "tags": [],
        "img": "https://s.koubei.com/rTr5xDyFGy7yJN7NeM7ctWpArRjf5xQc.jpg"
      },
      {
        "id": "p204",
        "name": "湖南津市特色麻辣牛肉",
        "price": 39.0,
        "desc": "湖南特色风味，干香麻辣，嚼劲十足！",
        "tags": [],
        "img": "https://s.koubei.com/zTQzZSkMXP4YyayEijfb2wdbN3FCpbsj.jpg",
        "opts": [
          {
            "g": "辣度",
            "req": false,
            "opts": [
              "麻辣"
            ]
          }
        ],
        "specs": [
          {
            "name": "小份",
            "price": 39.0
          },
          {
            "name": "大份",
            "price": 78.0
          }
        ]
      }
    ]
  },
  {
    "id": "cat3",
    "name": "手工鲜制饮品",
    "items": [
      {
        "id": "p301",
        "name": "芒果珍珠冰豆花",
        "price": 8.0,
        "desc": "豆花嫩滑、芒果香甜、珍珠Q弹，果香与焦糖交织，层次丰富。",
        "tags": [],
        "img": "https://s.koubei.com/f6zQxYXCKjei7ep8awaWm8CDhzXNdMw4.jpg"
      },
      {
        "id": "p302",
        "name": "木薯糖水",
        "price": 6.0,
        "desc": "木薯软糯，糖水温润",
        "tags": [],
        "img": "https://s.koubei.com/EfXPS2mm8EaMnARbSpXnjC2ExFwzYaM5.jpg"
      },
      {
        "id": "p303",
        "name": "桂花蜜豆花",
        "price": 6.0,
        "desc": "豆花滑嫩，桂花蜜清甜",
        "tags": [],
        "img": "https://s.koubei.com/SH7NEyhzH6W4w3XXD6RZnHSjm7z7HiBi.jpg"
      },
      {
        "id": "p304",
        "name": "桂花米酒汤圆",
        "price": 6.0,
        "desc": "老味道桂花米酒汤圆，酒酿鲜醇，圆子软糯，桂香淡淡的，甜而不齁",
        "tags": [],
        "img": "https://s.koubei.com/M53mDEJn784r7kb2sDmiNWnwedaspfNP.jpg"
      }
    ]
  },
  {
    "id": "cat4",
    "name": "鲜鱼汤-现炒码子粉",
    "items": [
      {
        "id": "p401",
        "name": "汉川特色-财鱼粉",
        "price": 32.0,
        "desc": "鲜财鱼脆嫩，重爆炒锅气，红汤浓郁，微辣回甜，香气浓烈，重口鲜香。",
        "tags": [],
        "img": "https://s.koubei.com/yDcYKzy6QjFWcfEYbixXTETKnPy8RGJ5.jpg",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "圆粉",
              "宽粉",
              "手工大宽粉",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "碱面",
              "豆丝",
              "湖南宽粉",
              "扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": false,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】品牌餐盒",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p402",
        "name": "汉川特色-鳝鱼粉",
        "price": 48.0,
        "desc": "鳝鱼滑弹鲜嫩，重爆炒锅气，红汤浓郁，微辣回甜，香气浓烈，重口鲜香。",
        "tags": [],
        "img": "https://s.koubei.com/zhKwkjKF7JCC84y5EhJAkaGzTsXb5GRy.jpg",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "圆粉",
              "宽粉",
              "手工大宽粉",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "碱面",
              "豆丝",
              "湖南宽粉",
              "扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": false,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】品牌餐盒",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p403",
        "name": "鲜鱼汤-鲜牛肉粉",
        "price": 36.0,
        "desc": "鲜牛肉与鱼汤融合，肉香与鲜甜交织，汤底鲜辣",
        "tags": [],
        "img": "https://s.koubei.com/DkbCSYhAwTHJXpayw2TfWmhjFTnSEMCD.jpg",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "圆粉",
              "宽粉",
              "手工大宽粉",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "碱面",
              "豆丝",
              "湖南宽粉",
              "扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": false,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】品牌餐盒",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          },
          {
            "g": "鲜鱼汤码子粉",
            "req": true,
            "opts": [
              "新做法（不可免辣）"
            ]
          }
        ]
      },
      {
        "id": "p404",
        "name": "鲜鱼汤-瘦肉粉",
        "price": 22.0,
        "desc": "鲜猪肉配薄皮青椒，锅气十足，鲜甜微辣",
        "tags": [],
        "img": "https://s.koubei.com/CfYeFafw5TnixE2WnmRHcQpPynW8y4Kc.png",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "圆粉",
              "宽粉",
              "手工大宽粉",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "碱面",
              "豆丝",
              "湖南宽粉",
              "扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": false,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】品牌餐盒",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p405",
        "name": "鲜鱼汤-猪肝粉",
        "price": 20.0,
        "desc": "猪肝嫩滑，配青椒、洋葱、西红柿，汁浓味美",
        "tags": [],
        "img": "https://s.koubei.com/6MhNaNzWisnJ6fZtby7TrKb8Rjtswytz.jpg",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "圆粉",
              "宽粉",
              "手工大宽粉",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "碱面",
              "豆丝",
              "湖南宽粉",
              "扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": false,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】品牌餐盒",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p406",
        "name": "鲜鱼汤-瘦肉猪肝粉",
        "price": 24.0,
        "desc": "瘦肉与猪肝双拼，鲜嫩加倍，汤底鲜甜",
        "tags": [],
        "img": "https://s.koubei.com/jQ7CrN4E5xmNmTNwzCQsNSFz44wyzayz.jpg",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "圆粉",
              "宽粉",
              "手工大宽粉",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "碱面",
              "豆丝",
              "湖南宽粉",
              "扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": false,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】品牌餐盒",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ]
      }
    ]
  },
  {
    "id": "cat5",
    "name": "汉川掺汤",
    "items": [
      {
        "id": "p501",
        "name": "汉川掺汤",
        "price": 18.0,
        "desc": "汉川掺汤精选鲜嫩瘦肉，以滚烫高汤快速汆熟，佐以足量黑胡椒，肉香醇厚，暖胃饱腹。",
        "tags": [
          "做法原因不可续",
          "黑胡椒"
        ],
        "img": "https://s.koubei.com/NEasQ5Q7hrhKhpj4j465cXsY4XSBmcJJ.jpg",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "苕粉",
              "宽粉",
              "豆丝",
              "扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": false,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】品牌餐盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "瘦肉",
            "price": 18.0
          },
          {
            "name": "猪肝",
            "price": 18.0
          },
          {
            "name": "瘦肉猪肝",
            "price": 22.0
          },
          {
            "name": "财鱼",
            "price": 28.0
          },
          {
            "name": "鳝鱼",
            "price": 48.0
          },
          {
            "name": "鳝鱼财鱼",
            "price": 48.0
          },
          {
            "name": "牛肉",
            "price": 36.0
          }
        ]
      }
    ]
  },
  {
    "id": "cat6",
    "name": "广式鲜粥",
    "items": [
      {
        "id": "p601",
        "name": "广式鲜粥系列",
        "price": 12.0,
        "desc": "广式鲜粥以大米慢熬至绵滑，可选择搭配各种食材，鲜香温润、养胃易消化。",
        "tags": [
          "咸鲜"
        ],
        "img": "https://s.koubei.com/yZNk5xrWpi4Pmj2HCAPJnxJT3YSpXHPh.jpg",
        "specs": [
          {
            "name": "皮蛋瘦肉粥",
            "price": 12.0
          },
          {
            "name": "瘦肉粥",
            "price": 12.0
          },
          {
            "name": "猪肝粥",
            "price": 12.0
          },
          {
            "name": "瘦肉猪肝粥",
            "price": 16.0
          },
          {
            "name": "财鱼粥",
            "price": 18.0
          },
          {
            "name": "鳝鱼粥",
            "price": 38.0
          },
          {
            "name": "财鱼鳝鱼粥",
            "price": 32.0
          },
          {
            "name": "牛肉粥",
            "price": 28.0
          }
        ]
      }
    ]
  },
  {
    "id": "cat7",
    "name": "新品专区",
    "items": [
      {
        "id": "p701",
        "name": "秘制蟹钳面",
        "price": 48.0,
        "desc": "蟹钳鲜香麻辣，面条劲道",
        "tags": [
          "推荐",
          "做法原因不可续面"
        ],
        "img": "https://s.koubei.com/hk2pBSZwxcrX73eKGtD8M6E5tPCjhXsC.jpg"
      },
      {
        "id": "p702",
        "name": "炒菜-秘制蟹钳",
        "price": 78.0,
        "desc": "蟹钳鲜香麻辣，层次丰富",
        "tags": [],
        "img": "https://s.koubei.com/xTXc3rmNch23dYekDs5XHDE3DJpNPjhi.jpg"
      },
      {
        "id": "p703",
        "name": "砂锅番茄炖肉丸",
        "price": 36.0,
        "desc": "番茄慢熬红汤，酸甜浓郁，蛋花柔滑；手工肉丸吸饱鲜汁，弹嫩入味。",
        "tags": [
          "做法原因约15分钟出品"
        ],
        "img": "https://s.koubei.com/pj3HWxYA7a32WtePTMirtsTrw5rhabr4.jpg",
        "opts": [
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】炖锅1768打包盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p704",
        "name": "红烧蹄膀",
        "price": 19.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/RsnKzdydrGBWrZGWXaPEH3rkWxyeJMdd.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不要主食",
              "配-米饭",
              "配-圆粉",
              "配-宽粉",
              "配-湖南宽粉",
              "配-手工大宽粉",
              "配-手工细面",
              "配-手工扁面",
              "配-手工宽面",
              "配-碱面",
              "配-豆丝",
              "配-白米粥",
              "配-扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "盖码份量",
            "price": 19.0
          },
          {
            "name": "炒菜份量",
            "price": 38.0
          }
        ]
      },
      {
        "id": "p705",
        "name": "平菇肉片汤",
        "price": 3.0,
        "desc": "平菇鲜嫩、肉片滑嫩，汤清味鲜，暖胃爽口。",
        "tags": [
          "不可续"
        ],
        "img": "https://s.koubei.com/iHTtH7rKYD8NKYEJ2jZcC6DisREDimnb.jpg",
        "opts": [
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】品牌餐盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p706",
        "name": "白萝卜炖牛肋条",
        "price": 28.0,
        "desc": "清甜白萝卜配软嫩牛肉，汤鲜肉香，暖胃又下饭。",
        "tags": [],
        "img": "https://s.koubei.com/xPyirKCiwxi2W6aZ4HxWr8wWhCZJsPtZ.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不要主食",
              "配-米饭",
              "配-圆粉",
              "配-宽粉",
              "配-湖南宽粉",
              "配-手工大宽粉",
              "配-手工细面",
              "配-手工扁面",
              "配-手工宽面",
              "配-碱面",
              "配-豆丝",
              "配-白米粥",
              "配-扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "盖码份量",
            "price": 28.0
          },
          {
            "name": "炒菜份量",
            "price": 36.0
          }
        ]
      }
    ]
  },
  {
    "id": "cat8",
    "name": "鲜货现炒码子",
    "items": [
      {
        "id": "p801",
        "name": "爆炒筋头巴脑",
        "price": 29.0,
        "desc": "牛肉筋软糯耐嚼，外酥里嫩，鲜辣回甜，锅气十足",
        "tags": [],
        "img": "https://s.koubei.com/xdPzTZczsdwcRGYAbH3wE4WXXfDsxXPE.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不要主食",
              "配-米饭",
              "配-圆粉",
              "配-宽粉",
              "配-手工大宽粉",
              "配-手工细面",
              "配-手工扁面",
              "配-手工宽面",
              "配-碱面",
              "配-豆丝",
              "配-白米粥",
              "配-湖南宽粉",
              "配-扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "盖码份量",
            "price": 29.0
          },
          {
            "name": "炒菜份量",
            "price": 38.0
          }
        ]
      },
      {
        "id": "p802",
        "name": "鲜牛肉炒樟树港辣椒",
        "price": 32.0,
        "desc": "鲜牛肉滑嫩，辣椒鲜香下饭",
        "tags": [],
        "img": "https://s.koubei.com/iAZwBzhxPbTRNse7TNss7TF3bWNB4bmj.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不要主食",
              "配-米饭",
              "配-圆粉",
              "配-宽粉",
              "配-手工大宽粉",
              "配-手工细面",
              "配-手工扁面",
              "配-手工宽面",
              "配-碱面",
              "配-豆丝",
              "配-白米粥",
              "配-湖南宽粉",
              "配-扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "盖码份量",
            "price": 32.0
          },
          {
            "name": "炒菜份量",
            "price": 48.0
          }
        ]
      },
      {
        "id": "p803",
        "name": "酸萝卜炒牛肉",
        "price": 32.0,
        "desc": "酸萝卜酸甜脆爽，牛肉鲜嫩，酸辣开胃",
        "tags": [],
        "img": "https://s.koubei.com/6yzJ8PcwifHe5ZYjNzWsK6j3SAnwZYQX.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不要主食",
              "配-米饭",
              "配-圆粉",
              "配-宽粉",
              "配-手工大宽粉",
              "配-手工细面",
              "配-手工扁面",
              "配-手工宽面",
              "配-碱面",
              "配-豆丝",
              "配-白米粥",
              "配-湖南宽粉",
              "配-扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "盖码份量",
            "price": 32.0
          },
          {
            "name": "炒菜份量",
            "price": 48.0
          }
        ]
      },
      {
        "id": "p804",
        "name": "回锅牛肉",
        "price": 32.0,
        "desc": "牛肉卤香醇厚，软而不柴，酱香浓郁",
        "tags": [],
        "img": "https://s.koubei.com/Njyb323GMXnTXM8j6f7ESxQ4hcWJdskz.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不要主食",
              "配-米饭",
              "配-圆粉",
              "配-宽粉",
              "配-手工大宽粉",
              "配-手工细面",
              "配-手工扁面",
              "配-手工宽面",
              "配-碱面",
              "配-豆丝",
              "配-白米粥",
              "配-湖南宽粉",
              "配-扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "盖码份量",
            "price": 32.0
          },
          {
            "name": "炒菜份量",
            "price": 48.0
          }
        ]
      },
      {
        "id": "p805",
        "name": "卤藕炒回锅牛肉",
        "price": 32.0,
        "desc": "牛肉卤香醇厚，卤藕粉糯干香，微辣醇厚",
        "tags": [],
        "img": "https://s.koubei.com/AFKZz6HZjB6xinFYwkisetwAC6PBzxrQ.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不要主食",
              "配-米饭",
              "配-圆粉",
              "配-宽粉",
              "配-手工大宽粉",
              "配-手工细面",
              "配-手工扁面",
              "配-手工宽面",
              "配-碱面",
              "配-豆丝",
              "配-白米粥",
              "配-湖南宽粉",
              "配-扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "盖码份量",
            "price": 32.0
          },
          {
            "name": "炒菜份量",
            "price": 48.0
          }
        ]
      },
      {
        "id": "p806",
        "name": "四季豆炒牛肚",
        "price": 28.0,
        "desc": "牛肚脆爽，四季豆鲜嫩，中辣咸鲜",
        "tags": [],
        "img": "https://s.koubei.com/FaSasXYXKrpxMdAwAfYA6drY6QGfArNG.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不要主食",
              "配-米饭",
              "配-圆粉",
              "配-宽粉",
              "配-手工大宽粉",
              "配-手工细面",
              "配-手工扁面",
              "配-手工宽面",
              "配-碱面",
              "配-豆丝",
              "配-白米粥",
              "配-湖南宽粉",
              "配-扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "盖码份量",
            "price": 28.0
          },
          {
            "name": "炒菜份量",
            "price": 48.0
          }
        ]
      },
      {
        "id": "p807",
        "name": "樟树港辣椒炒肉",
        "price": 28.0,
        "desc": "下饭神器！！",
        "tags": [],
        "img": "https://s.koubei.com/dCAj7FEHn7JY8CPiNEiDzh8MyH3fxzGT.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不要主食",
              "配-米饭",
              "配-圆粉",
              "配-宽粉",
              "配-手工大宽粉",
              "配-手工细面",
              "配-手工扁面",
              "配-手工宽面",
              "配-碱面",
              "配-豆丝",
              "配-白米粥",
              "配-湖南宽粉",
              "配-扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "盖码份量",
            "price": 28.0
          },
          {
            "name": "炒菜份量",
            "price": 38.0
          }
        ]
      },
      {
        "id": "p808",
        "name": "香干回锅肉",
        "price": 24.0,
        "desc": "五花肉肥而不腻，香干外韧里嫩，酱香微辣",
        "tags": [],
        "img": "https://s.koubei.com/GCbEMjhJzEsBBzRJrzbMFbZ2paADzfWJ.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不要主食",
              "配-米饭",
              "配-圆粉",
              "配-宽粉",
              "配-手工大宽粉",
              "配-手工细面",
              "配-手工扁面",
              "配-手工宽面",
              "配-碱面",
              "配-豆丝",
              "配-白米粥",
              "配-湖南宽粉",
              "配-扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "盖码份量",
            "price": 24.0
          },
          {
            "name": "炒菜份量",
            "price": 32.0
          }
        ]
      },
      {
        "id": "p809",
        "name": "回锅肉炒卤藕",
        "price": 24.0,
        "desc": "五花肉肥而不腻，卤藕粉糯干香",
        "tags": [],
        "img": "https://s.koubei.com/WfebxPypzReQ2GXQC4GDdRYkcrP8BWp6.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不要主食",
              "配-米饭",
              "配-圆粉",
              "配-宽粉",
              "配-手工大宽粉",
              "配-手工细面",
              "配-手工扁面",
              "配-手工宽面",
              "配-碱面",
              "配-豆丝",
              "配-白米粥",
              "配-湖南宽粉",
              "配-扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "盖码份量",
            "price": 24.0
          },
          {
            "name": "炒菜份量",
            "price": 32.0
          }
        ]
      },
      {
        "id": "p810",
        "name": "爆炒顺风",
        "price": 24.0,
        "desc": "卤猪耳脆嫩爽口，咸鲜香辣，锅气十足",
        "tags": [],
        "img": "https://s.koubei.com/zCB6nfjHXtETnMnmb5Q2mZyCC4aYj8xw.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不要主食",
              "配-米饭",
              "配-圆粉",
              "配-宽粉",
              "配-手工大宽粉",
              "配-手工细面",
              "配-手工扁面",
              "配-手工宽面",
              "配-碱面",
              "配-豆丝",
              "配-白米粥",
              "配-湖南宽粉",
              "配-扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "盖码份量",
            "price": 24.0
          },
          {
            "name": "炒菜份量",
            "price": 38.0
          }
        ]
      },
      {
        "id": "p811",
        "name": "青椒肉丝",
        "price": 19.0,
        "desc": "青椒脆嫩，肉丝滑嫩，咸鲜微辣",
        "tags": [],
        "img": "https://s.koubei.com/4rpiFDSsaTrGGYt5Ze4tdk5HdYZttJ5Z.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不要主食",
              "配-米饭",
              "配-圆粉",
              "配-宽粉",
              "配-手工大宽粉",
              "配-手工细面",
              "配-手工扁面",
              "配-手工宽面",
              "配-碱面",
              "配-豆丝",
              "配-白米粥",
              "配-湖南宽粉",
              "配-扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "盖码份量",
            "price": 19.0
          },
          {
            "name": "炒菜份量",
            "price": 28.0
          }
        ]
      },
      {
        "id": "p812",
        "name": "洋葱肉丝",
        "price": 19.0,
        "desc": "洋葱清甜，肉丝嫩滑，咸鲜适口",
        "tags": [],
        "img": "https://s.koubei.com/3Q2cf6JSHsGmdrZ4mP8mrHFFmGdGWxdZ.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不要主食",
              "配-米饭",
              "配-圆粉",
              "配-宽粉",
              "配-手工大宽粉",
              "配-手工细面",
              "配-手工扁面",
              "配-手工宽面",
              "配-碱面",
              "配-豆丝",
              "配-白米粥",
              "配-湖南宽粉",
              "配-扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "盖码份量",
            "price": 19.0
          },
          {
            "name": "炒菜份量",
            "price": 28.0
          }
        ]
      },
      {
        "id": "p813",
        "name": "榨菜肉丝",
        "price": 19.0,
        "desc": "肉丝嫩滑，榨菜咸香，下饭神器",
        "tags": [],
        "img": "https://s.koubei.com/ZjAsdK3dTfcjnkXCGdH685nrNnC5ZaRJ.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不要主食",
              "配-米饭",
              "配-圆粉",
              "配-宽粉",
              "配-手工大宽粉",
              "配-手工细面",
              "配-手工扁面",
              "配-手工宽面",
              "配-碱面",
              "配-豆丝",
              "配-白米粥",
              "配-湖南宽粉",
              "配-扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "盖码份量",
            "price": 19.0
          },
          {
            "name": "炒菜份量",
            "price": 28.0
          }
        ]
      },
      {
        "id": "p814",
        "name": "鱼香肉丝",
        "price": 24.0,
        "desc": "肉丝滑嫩，酸甜微辣开胃",
        "tags": [],
        "img": "https://s.koubei.com/3FBRDcmr3Ki2xdAaEtStZTwSPWx7PnZc.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不要主食",
              "配-米饭",
              "配-圆粉",
              "配-宽粉",
              "配-手工大宽粉",
              "配-手工细面",
              "配-手工扁面",
              "配-手工宽面",
              "配-碱面",
              "配-豆丝",
              "配-白米粥",
              "配-湖南宽粉",
              "配-扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "盖码份量",
            "price": 24.0
          },
          {
            "name": "炒菜份量",
            "price": 38.0
          }
        ]
      },
      {
        "id": "p815",
        "name": "爆炒猪肝",
        "price": 19.0,
        "desc": "猪肝嫩滑，火候精准，咸鲜适口",
        "tags": [],
        "img": "https://s.koubei.com/5PZjTH5FrEtHYHnaCHjCMmSxJ8QQbaHp.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不要主食",
              "配-米饭",
              "配-圆粉",
              "配-宽粉",
              "配-手工大宽粉",
              "配-手工细面",
              "配-手工扁面",
              "配-手工宽面",
              "配-碱面",
              "配-豆丝",
              "配-白米粥",
              "配-湖南宽粉",
              "配-扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "盖码份量",
            "price": 19.0
          },
          {
            "name": "炒菜份量",
            "price": 28.0
          }
        ]
      },
      {
        "id": "p816",
        "name": "爆炒鳝鱼丝",
        "price": 48.0,
        "desc": "鳝鱼丝滑嫩，烧椒酱咸鲜微辣，复合风味",
        "tags": [],
        "img": "https://s.koubei.com/K4Tsh8ZCrhpWMj6kHNKi2jmaCbP2aQph.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不要主食",
              "配-米饭",
              "配-圆粉",
              "配-宽粉",
              "配-手工大宽粉",
              "配-手工细面",
              "配-手工扁面",
              "配-手工宽面",
              "配-碱面",
              "配-豆丝",
              "配-白米粥",
              "配-湖南宽粉",
              "配-扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "盖码份量",
            "price": 48.0
          },
          {
            "name": "炒菜份量",
            "price": 88.0
          }
        ]
      },
      {
        "id": "p817",
        "name": "爆炒财鱼片",
        "price": 28.0,
        "desc": "财鱼片脆嫩紧实，汁水浓郁，搭配黑胡椒碎，鲜甜入味",
        "tags": [],
        "img": "https://s.koubei.com/6ArpDft5BX7RNcpNsDzrjsKyAcjSWbSB.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不要主食",
              "配-米饭",
              "配-圆粉",
              "配-宽粉",
              "配-手工大宽粉",
              "配-手工细面",
              "配-手工扁面",
              "配-手工宽面",
              "配-碱面",
              "配-豆丝",
              "配-白米粥",
              "配-湖南宽粉",
              "配-扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "盖码份量",
            "price": 28.0
          },
          {
            "name": "炒菜份量",
            "price": 58.0
          }
        ]
      },
      {
        "id": "p818",
        "name": "番茄炒蛋",
        "price": 15.0,
        "desc": "酸甜开胃，经典家常味",
        "tags": [
          "咸甜"
        ],
        "img": "https://s.koubei.com/3QiH37pYbaiFtyNdxjYsPaTYHBBCbR7D.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不要主食",
              "配-米饭",
              "配-圆粉",
              "配-宽粉",
              "配-手工大宽粉",
              "配-手工细面",
              "配-手工扁面",
              "配-手工宽面",
              "配-碱面",
              "配-豆丝",
              "配-白米粥",
              "配-湖南宽粉",
              "配-扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "盖码份量",
            "price": 15.0
          },
          {
            "name": "炒菜份量",
            "price": 22.0
          }
        ]
      },
      {
        "id": "p819",
        "name": "火爆鸡胗",
        "price": 22.0,
        "desc": "鸡胗脆嫩，猛火爆炒，香辣下饭",
        "tags": [],
        "img": "https://s.koubei.com/M4rE4A6rBbQww3FR8KTwHmSphm6yy4bh.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不要主食",
              "配-米饭",
              "配-圆粉",
              "配-宽粉",
              "配-手工大宽粉",
              "配-手工细面",
              "配-手工扁面",
              "配-手工宽面",
              "配-碱面",
              "配-豆丝",
              "配-白米粥",
              "配-湖南宽粉",
              "配-扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "盖码份量",
            "price": 22.0
          },
          {
            "name": "炒菜份量",
            "price": 32.0
          }
        ]
      },
      {
        "id": "p820",
        "name": "小炒黄牛肉",
        "price": 35.0,
        "desc": "鲜牛肉猛火爆炒，滑嫩鲜辣，油亮诱人。",
        "tags": [],
        "img": "https://s.koubei.com/6TmpycEnSJNEe6ZKSB2hNyX4aNHey8db.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不要主食",
              "配-米饭",
              "配-圆粉",
              "配-宽粉",
              "配-手工大宽粉",
              "配-手工细面",
              "配-手工扁面",
              "配-手工宽面",
              "配-碱面",
              "配-豆丝",
              "配-白米粥",
              "配-湖南宽粉",
              "配-扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "盖码份量",
            "price": 35.0
          },
          {
            "name": "炒菜份量",
            "price": 48.0
          }
        ]
      },
      {
        "id": "p821",
        "name": "洋葱炒蛋",
        "price": 15.0,
        "desc": "",
        "tags": [
          "咸鲜"
        ],
        "img": "https://s.koubei.com/n4zfPGPGW4a32En2nY4d4SsiM4ikeATt.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不要主食",
              "配-米饭",
              "配-圆粉",
              "配-宽粉",
              "配-手工大宽粉",
              "配-手工细面",
              "配-手工扁面",
              "配-手工宽面",
              "配-碱面",
              "配-豆丝",
              "配-白米粥",
              "配-湖南宽粉",
              "配-扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "盖码份量",
            "price": 15.0
          },
          {
            "name": "炒菜份量",
            "price": 22.0
          }
        ]
      },
      {
        "id": "p822",
        "name": "丝瓜炒蛋",
        "price": 18.0,
        "desc": "",
        "tags": [
          "咸甜"
        ],
        "img": "https://s.koubei.com/jwx73p8EDt34ChNfW8dyXeEZYwHSBTXk.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不要主食",
              "配-米饭",
              "配-圆粉",
              "配-宽粉",
              "配-手工大宽粉",
              "配-手工细面",
              "配-手工扁面",
              "配-手工宽面",
              "配-碱面",
              "配-豆丝",
              "配-白米粥",
              "配-湖南宽粉",
              "配-扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "盖码份量",
            "price": 18.0
          },
          {
            "name": "炒菜份量",
            "price": 24.0
          }
        ]
      },
      {
        "id": "p823",
        "name": "素炒小白菜",
        "price": 9.0,
        "desc": "整体少油入口脆爽，清爽解腻，水润回甘，软嫩不软烂！",
        "tags": [],
        "img": "https://s.koubei.com/fZGCDRM4rDbfx3yQcwyE2NjkikFkN6SM.jpg",
        "opts": [
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "盖码份量",
            "price": 9.0
          },
          {
            "name": "炒菜份量",
            "price": 16.0
          }
        ]
      }
    ]
  },
  {
    "id": "cat9",
    "name": "特色炖锅",
    "items": [
      {
        "id": "p901",
        "name": "特色牛杂炖粉",
        "price": 68.0,
        "desc": "新鲜牛肉、牛肚、牛筋、牛肠、牛黄喉配秘制锅底，细火慢炖，辣滚香鲜，汤醇味厚；汤汁通过高温融入米粉，入口软烂适中，咸淡适宜。",
        "tags": [],
        "img": "https://s.koubei.com/neA5isTiMPnST5Zbj5n5K2JPEK8WnGXY.jpg",
        "opts": [
          {
            "g": "辣度",
            "req": true,
            "opts": [
              "微辣",
              "清汤不辣",
              "中辣",
              "特辣"
            ]
          },
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "圆粉",
              "宽粉",
              "手工大宽粉",
              "扁粉"
            ]
          }
        ],
        "specs": [
          {
            "name": "小锅",
            "price": 68.0
          },
          {
            "name": "中锅",
            "price": 98.0
          },
          {
            "name": "大锅",
            "price": 128.0
          }
        ]
      },
      {
        "id": "p902",
        "name": "光头粉（只能搭配炖锅下单）",
        "price": 5.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/4AWeHpYTDT3hw2waxPzz7ch3QH2r7j5n.jpg",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "圆粉",
              "宽粉",
              "湖南宽粉",
              "扁粉"
            ]
          }
        ]
      },
      {
        "id": "p903",
        "name": "时蔬拼盘（配炖锅下单）",
        "price": 18.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/3NjFfj23CiACrN83ijYyQ7iNTMM8H7fM.jpg"
      }
    ]
  },
  {
    "id": "cat10",
    "name": "鲜货炭火烧烤",
    "items": [
      {
        "id": "p1001",
        "name": "原味羊肉串",
        "price": 5.0,
        "desc": "羊肉鲜嫩，炭烤香气浓郁，默认微辣",
        "tags": [],
        "img": "https://s.koubei.com/4t2QAciKHByY6cQ2wZp2nwy2Z72zefNZ.jpg",
        "opts": [
          {
            "g": "辣度",
            "req": true,
            "opts": [
              "微辣",
              "中辣",
              "特辣",
              "不要辣"
            ]
          },
          {
            "g": "小吃做法",
            "req": false,
            "opts": [
              "打包带走",
              "不要葱花"
            ]
          }
        ]
      },
      {
        "id": "p1002",
        "name": "原味猪肉串",
        "price": 3.0,
        "desc": "鲜肉现穿，清淡原味，肉汁丰富。",
        "tags": [],
        "img": "https://s.koubei.com/enk8Em3QXbmGSzsZ6fBHhTYmpFCazEnW.jpg",
        "opts": [
          {
            "g": "辣度",
            "req": true,
            "opts": [
              "微辣",
              "中辣",
              "特辣",
              "不要辣"
            ]
          },
          {
            "g": "小吃做法",
            "req": false,
            "opts": [
              "打包带走",
              "不要葱花"
            ]
          }
        ]
      },
      {
        "id": "p1003",
        "name": "牛肋条",
        "price": 6.0,
        "desc": "鲜牛肉肉质细嫩带微脆筋，烤后多汁，孜然肉香浓郁",
        "tags": [],
        "img": "https://s.koubei.com/HhxjXThFJby3CXQZP333dfTaACCfJQEm.jpg",
        "opts": [
          {
            "g": "辣度",
            "req": true,
            "opts": [
              "微辣",
              "中辣",
              "特辣",
              "不要辣"
            ]
          },
          {
            "g": "小吃做法",
            "req": false,
            "opts": [
              "打包带走",
              "不要葱花"
            ]
          }
        ]
      },
      {
        "id": "p1004",
        "name": "猪脆骨",
        "price": 3.0,
        "desc": "脆嫩有嚼劲，烧烤风味十足，默认微辣",
        "tags": [],
        "img": "https://s.koubei.com/dGR464P8aSYmWj8SCnG6Wa8KywNTCMja.jpg",
        "opts": [
          {
            "g": "辣度",
            "req": true,
            "opts": [
              "微辣",
              "中辣",
              "特辣",
              "不要辣"
            ]
          },
          {
            "g": "小吃做法",
            "req": false,
            "opts": [
              "打包带走",
              "不要葱花"
            ]
          }
        ]
      },
      {
        "id": "p1005",
        "name": "王中王火腿肠",
        "price": 5.0,
        "desc": "外焦里嫩，经典烧烤风味，默认微辣",
        "tags": [],
        "img": "https://s.koubei.com/32AG4QDEFyXbEimAcnxdydpJH3WctSB6.jpg",
        "opts": [
          {
            "g": "辣度",
            "req": true,
            "opts": [
              "微辣",
              "中辣",
              "特辣",
              "不要辣"
            ]
          },
          {
            "g": "小吃做法",
            "req": false,
            "opts": [
              "打包带走",
              "不要葱花"
            ]
          }
        ]
      },
      {
        "id": "p1006",
        "name": "烤凤爪",
        "price": 12.0,
        "desc": "先卤后烤，软糯Q弹",
        "tags": [],
        "img": "https://s.koubei.com/kRmy2JH3GQefnADjPxs48K4RXn32yCbn.jpg",
        "opts": [
          {
            "g": "辣度",
            "req": true,
            "opts": [
              "微辣",
              "中辣",
              "特辣",
              "不要辣"
            ]
          },
          {
            "g": "小吃做法",
            "req": false,
            "opts": [
              "打包带走",
              "不要葱花"
            ]
          }
        ]
      },
      {
        "id": "p1007",
        "name": "西葫芦",
        "price": 4.0,
        "desc": "烤后清甜，软嫩多汁，清爽解腻",
        "tags": [],
        "img": "https://s.koubei.com/x7xGYYpTKzkAyWjF43kaayFpbFKBCwHS.jpg",
        "opts": [
          {
            "g": "辣度",
            "req": true,
            "opts": [
              "微辣",
              "中辣",
              "特辣",
              "不要辣"
            ]
          },
          {
            "g": "小吃做法",
            "req": false,
            "opts": [
              "打包带走",
              "不要葱花"
            ]
          }
        ]
      },
      {
        "id": "p1008",
        "name": "烤鸡胗",
        "price": 3.0,
        "desc": "炭火现烤鸡胗，外皮微焦内里脆弹，满满粗孜然香，越嚼越上头。",
        "tags": [],
        "img": "https://s.koubei.com/GJFpw43FszdnFr8HjaDaznSMjRe23njz.jpg",
        "opts": [
          {
            "g": "辣度",
            "req": true,
            "opts": [
              "微辣",
              "中辣",
              "特辣",
              "不要辣"
            ]
          },
          {
            "g": "小吃做法",
            "req": false,
            "opts": [
              "打包带走",
              "不要葱花"
            ]
          }
        ]
      },
      {
        "id": "p1009",
        "name": "烤黄喉",
        "price": 5.0,
        "desc": "黄喉脆嫩爽口，炭火烤制，默认微辣",
        "tags": [],
        "img": "https://s.koubei.com/3EEZWQaNXjJZzSnjaG7aYEh8x5ASxBFi.jpg",
        "opts": [
          {
            "g": "辣度",
            "req": true,
            "opts": [
              "微辣",
              "中辣",
              "特辣",
              "不要辣"
            ]
          },
          {
            "g": "小吃做法",
            "req": false,
            "opts": [
              "打包带走",
              "不要葱花"
            ]
          }
        ]
      },
      {
        "id": "p1010",
        "name": "糯筋子",
        "price": 3.0,
        "desc": "猪肉鲜香嫩口，现点现烤，默认微辣",
        "tags": [],
        "img": "https://s.koubei.com/2i3wECKKkN6WpnQSaHtGDrkkwMMtFpGF.jpg",
        "opts": [
          {
            "g": "辣度",
            "req": true,
            "opts": [
              "微辣",
              "中辣",
              "特辣",
              "不要辣"
            ]
          },
          {
            "g": "小吃做法",
            "req": false,
            "opts": [
              "打包带走",
              "不要葱花"
            ]
          }
        ]
      },
      {
        "id": "p1011",
        "name": "红糖烧饼",
        "price": 8.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/c8QHKCKQsDrhNmRnbaBctimdZNT8Z3Dp.jpg",
        "opts": [
          {
            "g": "辣度",
            "req": true,
            "opts": [
              "微辣",
              "中辣",
              "特辣",
              "不要辣"
            ]
          },
          {
            "g": "小吃做法",
            "req": false,
            "opts": [
              "打包带走",
              "不要葱花"
            ]
          }
        ]
      },
      {
        "id": "p1012",
        "name": "香蕉吐司",
        "price": 8.0,
        "desc": "外酥内软，老少皆宜，微甜",
        "tags": [
          "咸甜"
        ],
        "img": "https://s.koubei.com/Y7hS3K2SBKB8KkSAaDyxyZCxFetexpJh.jpg",
        "opts": [
          {
            "g": "小吃做法",
            "req": false,
            "opts": [
              "打包带走",
              "不要葱花"
            ]
          }
        ]
      }
    ]
  },
  {
    "id": "cat11",
    "name": "时令鲜汤",
    "items": [
      {
        "id": "p1101",
        "name": "财鱼皮蛋豆腐汤",
        "price": 28.0,
        "desc": "财鱼鲜嫩，皮蛋提香，汤清味鲜",
        "tags": [
          "咸鲜"
        ],
        "img": "https://s.koubei.com/4bhXSsdGzZd5AYJCmimKGCXMEDCZm6sP.jpg",
        "opts": [
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】品牌餐盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p1102",
        "name": "番茄蛋汤",
        "price": 18.0,
        "desc": "酸甜开胃，经典家常汤品",
        "tags": [
          "咸鲜",
          "微酸"
        ],
        "img": "https://s.koubei.com/CMncDeY8YD7nRK74Jw8wMA66MXN2Y6sn.jpg",
        "opts": [
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】品牌餐盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p1103",
        "name": "番茄肉片蛋汤",
        "price": 24.0,
        "desc": "番茄酸甜，肉片蛋花鲜嫩",
        "tags": [
          "咸鲜",
          "微酸"
        ],
        "img": "https://s.koubei.com/rBRK8ZkY5QKRxzKSDKfJetCJt5nFhbQF.jpg",
        "opts": [
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】品牌餐盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p1104",
        "name": "排骨藕汤/单人份",
        "price": 12.0,
        "desc": "排骨软烂，莲藕软糯，汤鲜味浓",
        "tags": [
          "咸鲜",
          "黑胡椒"
        ],
        "img": "https://s.koubei.com/PTWQTRnQFDWsjJ6s55yMEmPnATSHmYAf.jpg",
        "opts": [
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】品牌餐盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p1105",
        "name": "标准份-排骨藕汤",
        "price": 22.0,
        "desc": "适合2~3人分享，汤鲜味浓，滋补暖身",
        "tags": [
          "咸鲜",
          "黑胡椒"
        ],
        "img": "https://s.koubei.com/BGJ4RR2P72mHnnCR2azCQySMjWPjcdM8.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不要主食",
              "配-米饭",
              "配-圆粉",
              "配-扁粉",
              "配-宽粉",
              "配-湖南宽粉",
              "配-手工大宽粉",
              "配-手工细面",
              "配-手工扁面",
              "配-手工宽面",
              "配-碱面",
              "配-豆丝"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】品牌餐盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p1106",
        "name": "丝瓜汤系列",
        "price": 18.0,
        "desc": "丝瓜清甜软嫩，牛肉滑嫩鲜香，汤清味鲜，清爽不腻。",
        "tags": [
          "咸鲜"
        ],
        "img": "https://s.koubei.com/NdSYMSExQXJZA72HTNX5keFj3hmHtNDy.jpg",
        "opts": [
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】品牌餐盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "丝瓜牛肉",
            "price": 32.0
          },
          {
            "name": "丝瓜鳝鱼",
            "price": 68.0
          },
          {
            "name": "丝瓜肉片蛋汤",
            "price": 28.0
          },
          {
            "name": "丝瓜蛋汤",
            "price": 18.0
          }
        ]
      }
    ]
  },
  {
    "id": "cat12",
    "name": "卤菜凉菜",
    "items": [
      {
        "id": "p1201",
        "name": "卤鸭脖/根",
        "price": 8.0,
        "desc": "酱香醇厚，咸甜微辣，鸭脖紧实耐啃，油润入味，解馋过瘾。",
        "tags": [],
        "img": "https://s.koubei.com/ccYx54QxhEj5MSCiywaaxC5Thf652caB.jpg",
        "opts": [
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】码子方盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p1202",
        "name": "凉拌卤牛肉",
        "price": 24.0,
        "desc": "卤牛肉醇香，凉拌鲜辣爽口",
        "tags": [],
        "img": "https://s.koubei.com/2NAQFTREkk7xejRyRhGwJsFPTswPArzn.jpg",
        "opts": [
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】码子方盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p1203",
        "name": "凉拌卤顺风",
        "price": 18.0,
        "desc": "卤猪耳脆嫩，凉拌香辣爽口",
        "tags": [],
        "img": "https://s.koubei.com/ZMWKXd2BYpcW2PZ8Sedne3p44NniQMJe.jpg",
        "opts": [
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】码子方盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p1204",
        "name": "皮蛋拌豆腐",
        "price": 12.0,
        "desc": "豆腐滑嫩，皮蛋鲜香开胃",
        "tags": [],
        "img": "https://s.koubei.com/GtwdBiJTZGmicCWkGQYpGTkb6mkwHAs3.jpg",
        "opts": [
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】码子方盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p1205",
        "name": "凉拌黄瓜",
        "price": 8.0,
        "desc": "黄瓜爽脆，蒜香酸辣开胃",
        "tags": [],
        "img": "https://s.koubei.com/EMxtB4inbtwZM7jJaBtah6BC6Mwyi7SF.jpg",
        "opts": [
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】码子方盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p1206",
        "name": "麻辣蛤蜊",
        "price": 12.0,
        "desc": "蛤蜊鲜嫩，鲜甜入味",
        "tags": [],
        "img": "https://s.koubei.com/CfZAQXcCCa8CxYKFGrASzsBjjSHnXEYW.jpg",
        "opts": [
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】码子方盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p1207",
        "name": "凉拌卤藕",
        "price": 12.0,
        "desc": "卤藕粉糯脆爽，凉拌鲜香",
        "tags": [],
        "img": "https://s.koubei.com/ejQpN27aSBJjtfBPEBTwNeHTRkaSsGdi.jpg",
        "opts": [
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】码子方盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p1208",
        "name": "凉拌毛豆",
        "price": 12.0,
        "desc": "毛豆鲜嫩，咸香微辣",
        "tags": [],
        "img": "https://s.koubei.com/W3nYeBhAPt6H5BaaKErGiyHzerf46A3p.jpg",
        "opts": [
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】码子方盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p1209",
        "name": "凉拌卤拼盘",
        "price": 28.0,
        "desc": "多款卤味拼盘，咸香丰富",
        "tags": [],
        "img": "https://s.koubei.com/Eye3eT7Dn66eaRCT4Tb8A84khGMEYxs7.jpg",
        "opts": [
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】码子方盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p1210",
        "name": "养生黑木耳",
        "price": 8.0,
        "desc": "木耳脆爽，酸辣清新开胃",
        "tags": [],
        "img": "https://s.koubei.com/Gs4fJp8heXnWRyeMddpjH4sPhassmEGS.jpg",
        "opts": [
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】码子方盒",
              "其他需求请备注"
            ]
          }
        ]
      }
    ]
  },
  {
    "id": "cat13",
    "name": "炒粉炒面炒饭炒豆丝",
    "items": [
      {
        "id": "p1301",
        "name": "（炒粉面饭）鲜牛肉",
        "price": 32.0,
        "desc": "鲜牛肉现炒，锅气十足，微辣入味；默认微辣",
        "tags": [
          "微辣",
          "做法原因不可续加"
        ],
        "img": "https://s.koubei.com/aJt8scATZQiTXHAeeezpfWzM244ijJS8.jpg",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "宽粉",
              "碱面",
              "粉面混炒",
              "炒饭",
              "豆丝"
            ]
          },
          {
            "g": "炒粉面饭做法",
            "req": false,
            "opts": [
              "额外加一个鸡蛋炒",
              "不要葱花",
              "不要豆芽",
              "不要酸豆角",
              "不要辣"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】品牌餐盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p1302",
        "name": "（炒粉面饭）卤牛肉",
        "price": 32.0,
        "desc": "卤牛肉干香，炒制后风味浓郁；默认微辣",
        "tags": [
          "微辣",
          "做法原因不可续加"
        ],
        "img": "https://s.koubei.com/bf3BRJnZSMxzTkWk8b77MrYM5JzCyhP3.jpg",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "宽粉",
              "碱面",
              "粉面混炒",
              "炒饭",
              "豆丝"
            ]
          },
          {
            "g": "炒粉面饭做法",
            "req": false,
            "opts": [
              "额外加一个鸡蛋炒",
              "不要葱花",
              "不要豆芽",
              "不要酸豆角",
              "不要辣"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】品牌餐盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p1303",
        "name": "（炒粉面饭）鲜猪肉",
        "price": 22.0,
        "desc": "猪肉鲜嫩，搭配粉面饭，经典炒制；默认微辣",
        "tags": [
          "微辣",
          "做法原因不可续加"
        ],
        "img": "https://s.koubei.com/xirnhkYkwCZRJm4AiS5CDHryP7KcXWcC.jpg",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "宽粉",
              "碱面",
              "粉面混炒",
              "炒饭",
              "豆丝"
            ]
          },
          {
            "g": "炒粉面饭做法",
            "req": false,
            "opts": [
              "额外加一个鸡蛋炒",
              "不要葱花",
              "不要豆芽",
              "不要酸豆角",
              "不要辣"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】品牌餐盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p1304",
        "name": "（炒粉面饭）鸡蛋火腿",
        "price": 18.0,
        "desc": "鸡蛋香嫩，火腿咸香，经典搭配；默认微辣",
        "tags": [
          "微辣",
          "做法原因不可续加"
        ],
        "img": "https://s.koubei.com/5inwRmYiB5663w2TPzhKKiY8bGYBki32.jpg",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "宽粉",
              "碱面",
              "粉面混炒",
              "炒饭",
              "豆丝"
            ]
          },
          {
            "g": "炒粉面饭做法",
            "req": false,
            "opts": [
              "额外加一个鸡蛋炒",
              "不要葱花",
              "不要豆芽",
              "不要酸豆角",
              "不要辣"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】品牌餐盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p1305",
        "name": "（炒粉面饭）卤顺风",
        "price": 24.0,
        "desc": "卤猪耳脆嫩；默认微辣",
        "tags": [
          "微辣",
          "做法原因不可续加"
        ],
        "img": "https://s.koubei.com/8t5mZYPsmFmHwpfasKsNZGeRx363r6Xw.jpg",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "宽粉",
              "碱面",
              "粉面混炒",
              "炒饭",
              "豆丝"
            ]
          },
          {
            "g": "炒粉面饭做法",
            "req": false,
            "opts": [
              "额外加一个鸡蛋炒",
              "不要葱花",
              "不要豆芽",
              "不要酸豆角",
              "不要辣"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】牛皮纸碗",
              "【打包】品牌餐盒",
              "其他需求请备注"
            ]
          }
        ]
      }
    ]
  },
  {
    "id": "cat14",
    "name": "主食",
    "items": [
      {
        "id": "p1401",
        "name": "五常大米饭",
        "price": 3.0,
        "desc": "五常大米蒸制，米香软糯",
        "tags": [],
        "img": "https://s.koubei.com/kh22sxhaz6N5PkSJzGQ5JjbXjbHj5tHt.jpg"
      }
    ]
  },
  {
    "id": "cat15",
    "name": "酒水饮料",
    "items": [
      {
        "id": "p1501",
        "name": "百岁山",
        "price": 3.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/GhFGsXA2yM6haCyXAWAGenYBRmbbM2Bf.jpg",
        "opts": [
          {
            "g": "温度",
            "req": false,
            "opts": [
              "常温",
              "冰"
            ]
          }
        ]
      },
      {
        "id": "p1502",
        "name": "百事可乐",
        "price": 3.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/DK5Km8DsnbjpFWMwPx6wHPpcxQbMdNkA.jpg",
        "opts": [
          {
            "g": "温度",
            "req": false,
            "opts": [
              "常温",
              "冰"
            ]
          }
        ]
      },
      {
        "id": "p1503",
        "name": "雪碧",
        "price": 3.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/487ZpdBtAK3waMksE46fG6sCetRnzMMm.jpg",
        "opts": [
          {
            "g": "温度",
            "req": false,
            "opts": [
              "常温",
              "冰"
            ]
          }
        ]
      },
      {
        "id": "p1504",
        "name": "芬达",
        "price": 3.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/AcwYZ768AkPWsJydWYt3QWi4h8HREAjM.jpg",
        "opts": [
          {
            "g": "温度",
            "req": false,
            "opts": [
              "常温",
              "冰"
            ]
          }
        ]
      },
      {
        "id": "p1505",
        "name": "维他奶",
        "price": 4.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/zMWJa2EEFfSPK4QaYrhbsGyQdTasr5jf.jpg",
        "opts": [
          {
            "g": "温度",
            "req": false,
            "opts": [
              "常温",
              "冰"
            ]
          }
        ]
      },
      {
        "id": "p1506",
        "name": "娃哈哈（小）",
        "price": 3.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/7adjwFJmtPtkGHfy23QJTnK8cRXh2EZz.jpg",
        "opts": [
          {
            "g": "温度",
            "req": false,
            "opts": [
              "常温",
              "冰"
            ]
          }
        ]
      },
      {
        "id": "p1507",
        "name": "娃哈哈（大）",
        "price": 7.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/f6krPxYTzd5a8ATCJFmcRWKsbB8tMFpw.jpg",
        "opts": [
          {
            "g": "温度",
            "req": false,
            "opts": [
              "常温",
              "冰"
            ]
          }
        ]
      },
      {
        "id": "p1508",
        "name": "椰树椰汁（245ml）",
        "price": 6.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/5FXM2TefwMCpa4A2Ge5rTX34RMQHw3S6.jpg",
        "opts": [
          {
            "g": "温度",
            "req": false,
            "opts": [
              "常温",
              "冰"
            ]
          }
        ]
      },
      {
        "id": "p1509",
        "name": "（橙诺）大窑",
        "price": 6.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/ptJMYFBDQyDFaRxCmBtCFBW4sA2p2e7X.jpg",
        "opts": [
          {
            "g": "温度",
            "req": false,
            "opts": [
              "常温",
              "冰"
            ]
          }
        ]
      },
      {
        "id": "p1510",
        "name": "（荔爱）大窑",
        "price": 6.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/tNWcjAQhX25zEnpp5BjjWXcjfPXiS5ZQ.jpg",
        "opts": [
          {
            "g": "温度",
            "req": false,
            "opts": [
              "常温",
              "冰"
            ]
          }
        ]
      },
      {
        "id": "p1511",
        "name": "（嘉宾）大窑",
        "price": 6.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/SjaB32RJT7bRYk4jFJfAMQF8emWXJA8w.jpg",
        "opts": [
          {
            "g": "温度",
            "req": false,
            "opts": [
              "常温",
              "冰"
            ]
          }
        ]
      },
      {
        "id": "p1512",
        "name": "大可乐",
        "price": 12.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/wmMi2pzzZN4ZG62Fmdh2ixFnsGfKAdYz.jpg",
        "opts": [
          {
            "g": "温度",
            "req": false,
            "opts": [
              "常温",
              "冰"
            ]
          }
        ]
      },
      {
        "id": "p1513",
        "name": "大雪碧",
        "price": 12.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/jyjYeE7H4hNxT7HJxcrT82QZdfJTaQFC.jpeg",
        "opts": [
          {
            "g": "温度",
            "req": false,
            "opts": [
              "常温",
              "冰"
            ]
          }
        ]
      },
      {
        "id": "p1514",
        "name": "大椰汁",
        "price": 25.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/crYRjmdAwAfNNSxakTDAzeJ52DKS3izE.jpg",
        "opts": [
          {
            "g": "温度",
            "req": false,
            "opts": [
              "常温",
              "冰"
            ]
          }
        ]
      },
      {
        "id": "p1515",
        "name": "瓶装-雪花勇闯天涯",
        "price": 7.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/Y7am2BmHe2e6QnWRsR4hTHwkDnM4MXZr.jpg",
        "opts": [
          {
            "g": "温度",
            "req": false,
            "opts": [
              "常温",
              "冰"
            ]
          }
        ]
      },
      {
        "id": "p1516",
        "name": "草本毛铺",
        "price": 20.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/AHFnzpi6iABQRpZddZ457D2Jj5mcN577.jpg",
        "opts": [
          {
            "g": "温度",
            "req": false,
            "opts": [
              "常温",
              "冰"
            ]
          }
        ]
      },
      {
        "id": "p1517",
        "name": "瓶装-雪花纯生",
        "price": 10.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/8RHBy86ncbRpsEjdsPr4hS6nrhwZaAey.jpg",
        "opts": [
          {
            "g": "温度",
            "req": false,
            "opts": [
              "常温",
              "冰"
            ]
          }
        ]
      },
      {
        "id": "p1518",
        "name": "小劲酒",
        "price": 18.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/2WefBbXW6fF5jixPnRXm7BixzXEXKSdP.jpg",
        "opts": [
          {
            "g": "温度",
            "req": false,
            "opts": [
              "常温",
              "冰"
            ]
          }
        ]
      },
      {
        "id": "p1519",
        "name": "燕京U8啤酒",
        "price": 8.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/MbCK62cBkxDDktfAmKntPPnkhzNKxPNn.jpg",
        "opts": [
          {
            "g": "温度",
            "req": false,
            "opts": [
              "常温",
              "冰"
            ]
          }
        ]
      },
      {
        "id": "p1520",
        "name": "百威纯生啤酒",
        "price": 10.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/HTayYym3sZJYJFmpDafnxm7YS3BpT7jb.jpg",
        "opts": [
          {
            "g": "温度",
            "req": false,
            "opts": [
              "常温",
              "冰"
            ]
          }
        ]
      }
    ]
  }
];
