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
        "name": "招牌秘制牛肉粉",
        "price": 25.0,
        "desc": "文火慢卤切片，汤底醇厚，微辣鲜香",
        "tags": [
          "招牌必点"
        ],
        "img": "https://s.koubei.com/jmCAC3cHjTAdQG6xnsz6ScxXP4sD8meG.png",
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
              "【打包】干湿分离打包盒",
              "【打包】品牌餐盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p102",
        "name": "香辣牛蹄筋粉",
        "price": 28.0,
        "desc": "糯叽叽自带胶感，中辣过瘾",
        "tags": [],
        "img": "https://s.koubei.com/8mSmdrs72jjNtyiscPikWtZBa8cTNXSH.jpg",
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
              "加码-牛蹄筋"
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
              "【打包】干湿分离打包盒",
              "【打包】品牌餐盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p103",
        "name": "津市麻辣牛肉粉",
        "price": 28.0,
        "desc": "湖南特色风味，干香麻辣，嚼劲十足！",
        "tags": [
          "麻辣"
        ],
        "img": "https://s.koubei.com/W42HtDx4ADJnpFQwnMQsYHYJT7nxhC2Q.jpg",
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
              "【打包】干湿分离打包盒",
              "【打包】品牌餐盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p104",
        "name": "鲜切原味牛肉粉",
        "price": 25.0,
        "desc": "大片牛肉卤香浓郁，配牛骨原汤，清淡鲜美，可选不辣或微辣",
        "tags": [
          "清汤"
        ],
        "img": "https://s.koubei.com/hEikWbrck2bNiKDACRrQYAA22JH5MEzY.jpg",
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
              "牛清汤（不辣）",
              "红油汤（微辣）"
            ]
          },
          {
            "g": "温馨提示",
            "req": false,
            "opts": [
              "堂食",
              "【打包】品牌餐盒",
              "【打包】干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p105",
        "name": "炸酱粉面",
        "price": 12.0,
        "desc": "爽滑米粉裹满醇香肉酱，咸香浓郁、越拌越香。",
        "tags": [
          "做法原因不可续"
        ],
        "img": "https://s.koubei.com/GHnp3jEYNJHXi8eZKTP4ED78KrNGEspd.jpg",
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
              "【打包】品牌餐盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p106",
        "name": "热干系列",
        "price": 6.0,
        "desc": "默认微辣",
        "tags": [
          "做法原因不可续面"
        ],
        "img": "https://s.koubei.com/sXc7jKC4jk6emxs2SJyWRZCdWWmrx7mr.jpg",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "热干面",
              "热干-宽粉",
              "热干-圆粉",
              "热干-扁粉",
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
              "【打包】品牌餐盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "原味热干面",
            "price": 6.0
          },
          {
            "name": "秘制牛肉热干面",
            "price": 25.0
          }
        ]
      },
      {
        "id": "p107",
        "name": "原汤粉/面",
        "price": 6.0,
        "desc": "原汤清鲜，粉面爽滑",
        "tags": [],
        "img": "https://s.koubei.com/86rh6cFfXhnwESPNZrHFzeC5FKd58zhC.jpg",
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
              "红油汤（微辣）",
              "牛清汤（不辣）"
            ]
          },
          {
            "g": "温馨提示",
            "req": false,
            "opts": [
              "堂食",
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
        "img": "https://s.koubei.com/3ZpD3SeQWMJT3CCZHKChnhWeJ7pSMAyi.jpg"
      },
      {
        "id": "p202",
        "name": "卤鸡爪/只",
        "price": 6.0,
        "desc": "软糯脱骨，卤香甜辣浓郁",
        "tags": [],
        "img": "https://s.koubei.com/2hzr3b78rWhph5YQj4rNzs2GJHW3mHwd.png"
      },
      {
        "id": "p203",
        "name": "卤海带结2个",
        "price": 3.0,
        "desc": "卤香甜辣入味，爽脆有嚼劲",
        "tags": [],
        "img": "https://s.koubei.com/dJJpjnFPFmsR7BPkn4FbNM2hbfaaDScs.jpg"
      },
      {
        "id": "p204",
        "name": "湖南津市特色麻辣牛肉",
        "price": 39.0,
        "desc": "湖南特色风味，干香麻辣，嚼劲十足！",
        "tags": [],
        "img": "https://s.koubei.com/RNXP5WMN4ZKDdpt3pntEcnfyNHPSKDiT.png",
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
    "name": "鲜鱼汤-现炒码子粉",
    "items": [
      {
        "id": "p301",
        "name": "鲜鱼汤-现炒财鱼粉",
        "price": 32.0,
        "desc": "活养财鱼肉质紧实，鲜香浓郁！",
        "tags": [],
        "img": "https://s.koubei.com/bAn6pH2rTn8H5ptAn5ZKDCECjGMasYFG.png",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "圆粉",
              "扁粉",
              "宽粉",
              "湖南宽粉",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "碱面",
              "豆丝"
            ]
          },
          {
            "g": "辣度",
            "req": true,
            "opts": [
              "正常做法",
              "不要辣",
              "其他需求请备注"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】品牌餐盒",
              "【打包】干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p302",
        "name": "鲜鱼汤-现炒鳝鱼粉",
        "price": 48.0,
        "desc": "活养鳝鱼，口感脆嫩，配鲜鱼汤底，来一碗舒服！",
        "tags": [],
        "img": "https://s.koubei.com/JAeRkMnNxBpZAn6Rzz54BezNZWG7aZb4.png",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "圆粉",
              "扁粉",
              "宽粉",
              "湖南宽粉",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "碱面",
              "豆丝"
            ]
          },
          {
            "g": "辣度",
            "req": true,
            "opts": [
              "正常做法",
              "不要辣",
              "其他需求请备注"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】品牌餐盒",
              "【打包】干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p303",
        "name": "鲜鱼汤-现炒瘦肉粉",
        "price": 22.0,
        "desc": "鲜熬鱼汤配鲜嫩瘦肉，微辣提鲜更开胃",
        "tags": [],
        "img": "https://s.koubei.com/JipKBhB4Q8KWhEFTbchj8As8ed2F8S4e.png",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "圆粉",
              "扁粉",
              "宽粉",
              "湖南宽粉",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "碱面",
              "豆丝"
            ]
          },
          {
            "g": "辣度",
            "req": true,
            "opts": [
              "正常做法",
              "不要辣",
              "其他需求请备注"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】品牌餐盒",
              "【打包】干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ]
      },
      {
        "id": "p304",
        "name": "鲜鱼汤-现炒鲜牛肉粉",
        "price": 36.0,
        "desc": "鲜牛肉与鱼汤融合，肉香与鲜甜交织",
        "tags": [],
        "img": "https://s.koubei.com/J64WZN5BdYdibjcPnAMPaa8Y84kYABB6.jpg",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "圆粉",
              "扁粉",
              "宽粉",
              "湖南宽粉",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "碱面",
              "豆丝"
            ]
          },
          {
            "g": "辣度",
            "req": true,
            "opts": [
              "正常做法",
              "不要辣",
              "其他需求请备注"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
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
    "id": "cat4",
    "name": "汉川掺汤",
    "items": [
      {
        "id": "p401",
        "name": "汉川掺汤",
        "price": 18.0,
        "desc": "汉川掺汤精选鲜嫩瘦肉，以滚烫高汤快速汆熟，佐以足量黑胡椒，肉香醇厚，暖胃饱腹。",
        "tags": [
          "做法原因不可续",
          "黑胡椒"
        ],
        "img": "https://s.koubei.com/RpYskkRTxKTdfr4YkFMyZ4WMe6BpX8bG.jpg",
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
    "id": "cat5",
    "name": "广式鲜粥",
    "items": [
      {
        "id": "p501",
        "name": "广式鲜粥系列-单人份",
        "price": 12.0,
        "desc": "广式鲜粥以大米慢熬至绵滑，可选择搭配各种食材，鲜香温润、养胃易消化。",
        "tags": [
          "咸鲜"
        ],
        "img": "https://s.koubei.com/TzsHBtMskyY2tRnwmZwjjESjMyZZjCRA.jpg",
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
            "name": "财鱼粥",
            "price": 16.0
          },
          {
            "name": "鳝鱼粥",
            "price": 24.0
          },
          {
            "name": "财鱼鳝鱼粥",
            "price": 24.0
          },
          {
            "name": "牛肉粥",
            "price": 18.0
          }
        ]
      },
      {
        "id": "p502",
        "name": "广式鲜粥系列-2到3人份",
        "price": 24.0,
        "desc": "广式鲜粥以大米慢熬至绵滑，可选择搭配各种食材，鲜香温润、养胃易消化。",
        "tags": [
          "咸鲜"
        ],
        "img": "https://s.koubei.com/TzsHBtMskyY2tRnwmZwjjESjMyZZjCRA.jpg",
        "specs": [
          {
            "name": "皮蛋瘦肉粥",
            "price": 24.0
          },
          {
            "name": "瘦肉粥",
            "price": 24.0
          },
          {
            "name": "财鱼粥",
            "price": 32.0
          },
          {
            "name": "鳝鱼粥",
            "price": 48.0
          },
          {
            "name": "财鱼鳝鱼粥",
            "price": 48.0
          },
          {
            "name": "牛肉粥",
            "price": 36.0
          }
        ]
      }
    ]
  },
  {
    "id": "cat6",
    "name": "新品专区",
    "items": [
      {
        "id": "p601",
        "name": "泥蒿炒腊肠",
        "price": 28.0,
        "desc": "泥蒿清甜，腊肠咸香，咸鲜适口",
        "tags": [],
        "img": "https://s.koubei.com/iJGTmRazXpSCMTpYsHRCs8rYBNet2zr8.jpg",
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
              "配-豆丝",
              "配-白米粥"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】品牌餐盒",
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
        "id": "p602",
        "name": "爆炒阳雀湖甘蓝（农博会金奖）",
        "price": 12.0,
        "desc": "农博会金奖阳雀湖甘蓝，脆嫩清甜",
        "tags": [],
        "img": "https://s.koubei.com/5sP2Mdk83nrkzmEiR8ejXQ624XjspxQK.jpg",
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
            "name": "盖码份量",
            "price": 12.0
          },
          {
            "name": "炒菜份量",
            "price": 18.0
          }
        ]
      },
      {
        "id": "p603",
        "name": "秘制蟹钳面",
        "price": 48.0,
        "desc": "蟹钳鲜香麻辣，面条劲道",
        "tags": [
          "推荐",
          "做法原因不可续面"
        ],
        "img": "https://s.koubei.com/kS6Fjydjh3pkTJKY23CSrnFmXZsFM8BT.jpg"
      },
      {
        "id": "p604",
        "name": "炒菜-秘制蟹钳",
        "price": 78.0,
        "desc": "蟹钳鲜香麻辣，层次丰富",
        "tags": [],
        "img": "https://s.koubei.com/2ShdQjwxAamTPtTZjAppaAzWDs22mHit.jpg"
      }
    ]
  },
  {
    "id": "cat7",
    "name": "现炒盖码饭",
    "items": [
      {
        "id": "p701",
        "name": "猪肉类-盖码饭",
        "price": 19.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/cix4iZ38zFsKDTGHMGmWYn7jYwEQZb8H.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不要主食",
              "配-米饭"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】品牌餐盒",
              "【打包】干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "香干回锅肉",
            "price": 24.0
          },
          {
            "name": "樟树港辣椒炒肉",
            "price": 28.0
          },
          {
            "name": "青椒肉丝",
            "price": 19.0
          },
          {
            "name": "洋葱肉丝",
            "price": 19.0
          },
          {
            "name": "榨菜肉丝",
            "price": 19.0
          },
          {
            "name": "鱼香肉丝",
            "price": 24.0
          },
          {
            "name": "回锅肉炒卤藕",
            "price": 24.0
          },
          {
            "name": "爆炒顺风",
            "price": 24.0
          },
          {
            "name": "爆炒猪肝",
            "price": 19.0
          },
          {
            "name": "红烧蹄膀",
            "price": 19.0
          },
          {
            "name": "农家一碗香",
            "price": 22.0
          }
        ]
      },
      {
        "id": "p702",
        "name": "牛肉类-盖码饭",
        "price": 28.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/5QCTeBYR8Y5HwNHAyZMa7HC4mKBmcnyY.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不要主食",
              "配-米饭"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】品牌餐盒",
              "【打包】干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "酸萝卜炒牛肉",
            "price": 32.0
          },
          {
            "name": "卤藕炒回锅牛肉",
            "price": 32.0
          },
          {
            "name": "爆炒筋头巴脑",
            "price": 29.0
          },
          {
            "name": "鲜牛肉炒樟树港辣椒",
            "price": 32.0
          },
          {
            "name": "回锅牛肉",
            "price": 32.0
          },
          {
            "name": "四季豆炒牛肚",
            "price": 28.0
          },
          {
            "name": "小炒黄牛肉",
            "price": 35.0
          }
        ]
      },
      {
        "id": "p703",
        "name": "鱼鲜禽蛋类-盖码饭",
        "price": 15.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/ba7bSCeJGDhp4Esefh7KixiCftsyZWTy.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不要主食",
              "配-米饭"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】品牌餐盒",
              "【打包】干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "爆炒财鱼片",
            "price": 28.0
          },
          {
            "name": "爆炒鳝鱼丝",
            "price": 48.0
          },
          {
            "name": "番茄炒蛋",
            "price": 15.0
          },
          {
            "name": "火爆鸡胗",
            "price": 22.0
          },
          {
            "name": "丝瓜炒蛋",
            "price": 18.0
          },
          {
            "name": "青椒煎茄子",
            "price": 19.0
          },
          {
            "name": "肉沫茄子",
            "price": 19.0
          }
        ]
      }
    ]
  },
  {
    "id": "cat8",
    "name": "现炒码子粉",
    "items": [
      {
        "id": "p801",
        "name": "牛肉类-码子粉",
        "price": 28.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/cDzWwBsYteH8zA3EbKS4B8s8ba6tiREZ.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "配-圆粉",
              "配-宽粉",
              "配-手工大宽粉",
              "配-手工细面",
              "配-手工扁面",
              "配-手工宽面",
              "配-碱面",
              "配-豆丝",
              "配-湖南宽粉",
              "配-扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】品牌餐盒",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "酸萝卜炒牛肉",
            "price": 32.0
          },
          {
            "name": "卤藕炒回锅牛肉",
            "price": 32.0
          },
          {
            "name": "爆炒筋头巴脑",
            "price": 29.0
          },
          {
            "name": "鲜牛肉炒樟树港辣椒",
            "price": 32.0
          },
          {
            "name": "回锅牛肉",
            "price": 32.0
          },
          {
            "name": "四季豆炒牛肚",
            "price": 28.0
          },
          {
            "name": "小炒黄牛肉",
            "price": 35.0
          }
        ]
      },
      {
        "id": "p802",
        "name": "鱼鲜禽蛋类-码子粉",
        "price": 15.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/atbJMmF3ZPcZsCDiFyrpwEJ8ZfhbKa66.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "配-圆粉",
              "配-宽粉",
              "配-手工大宽粉",
              "配-手工细面",
              "配-手工扁面",
              "配-手工宽面",
              "配-碱面",
              "配-豆丝",
              "配-湖南宽粉",
              "配-扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】品牌餐盒",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "爆炒财鱼片",
            "price": 28.0
          },
          {
            "name": "爆炒鳝鱼丝",
            "price": 48.0
          },
          {
            "name": "番茄炒蛋",
            "price": 15.0
          },
          {
            "name": "火爆鸡胗",
            "price": 22.0
          },
          {
            "name": "丝瓜炒蛋",
            "price": 18.0
          },
          {
            "name": "青椒煎茄子",
            "price": 19.0
          },
          {
            "name": "肉沫茄子",
            "price": 19.0
          }
        ]
      },
      {
        "id": "p803",
        "name": "猪肉类-码子粉",
        "price": 19.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/S6zA3pmeFR8hkQZjS6ZMJsJf2hpKh4BY.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "配-圆粉",
              "配-宽粉",
              "配-手工大宽粉",
              "配-手工细面",
              "配-手工扁面",
              "配-手工宽面",
              "配-碱面",
              "配-豆丝",
              "配-湖南宽粉",
              "配-扁粉"
            ]
          },
          {
            "g": "温馨提示",
            "req": true,
            "opts": [
              "堂食",
              "【打包】品牌餐盒",
              "【打包】码子粉干湿分离打包盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "香干回锅肉",
            "price": 24.0
          },
          {
            "name": "樟树港辣椒炒肉",
            "price": 28.0
          },
          {
            "name": "青椒肉丝",
            "price": 19.0
          },
          {
            "name": "洋葱肉丝",
            "price": 19.0
          },
          {
            "name": "榨菜肉丝",
            "price": 19.0
          },
          {
            "name": "鱼香肉丝",
            "price": 24.0
          },
          {
            "name": "爆炒猪肝",
            "price": 19.0
          },
          {
            "name": "回锅肉炒卤藕",
            "price": 24.0
          },
          {
            "name": "爆炒顺风",
            "price": 24.0
          },
          {
            "name": "红烧蹄膀",
            "price": 19.0
          },
          {
            "name": "农家一碗香",
            "price": 22.0
          }
        ]
      }
    ]
  },
  {
    "id": "cat9",
    "name": "家常炒菜",
    "items": [
      {
        "id": "p901",
        "name": "时蔬类-炒菜份量",
        "price": 16.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/iYx8tkyhmbsaSDK8HrRkKZnwM4aKMDE3.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不含主食"
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
        ],
        "specs": [
          {
            "name": "爆炒阳雀湖甘蓝（农博会金奖）",
            "price": 18.0
          },
          {
            "name": "素炒小白菜",
            "price": 16.0
          },
          {
            "name": "清炒菜苔",
            "price": 18.0
          },
          {
            "name": "酸辣菜苔",
            "price": 18.0
          },
          {
            "name": "清炒泥蒿",
            "price": 28.0
          }
        ]
      },
      {
        "id": "p902",
        "name": "鱼鲜禽蛋类-炒菜",
        "price": 22.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/PGs5aG4na6HY3mbS3GBQT7APtTexytrN.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不含主食"
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
        ],
        "specs": [
          {
            "name": "爆炒财鱼片",
            "price": 58.0
          },
          {
            "name": "爆炒鳝鱼丝",
            "price": 88.0
          },
          {
            "name": "番茄炒蛋",
            "price": 22.0
          },
          {
            "name": "火爆鸡胗",
            "price": 32.0
          },
          {
            "name": "丝瓜炒蛋",
            "price": 24.0
          },
          {
            "name": "青椒煎茄子",
            "price": 28.0
          },
          {
            "name": "肉沫茄子",
            "price": 28.0
          }
        ]
      },
      {
        "id": "p903",
        "name": "牛肉类-炒菜",
        "price": 38.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/ffB88YA2cianiQjWpSeGZTWK4BAYNkjB.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不含主食"
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
        ],
        "specs": [
          {
            "name": "酸萝卜炒牛肉",
            "price": 48.0
          },
          {
            "name": "卤藕炒回锅牛肉",
            "price": 48.0
          },
          {
            "name": "爆炒筋头巴脑",
            "price": 38.0
          },
          {
            "name": "鲜牛肉炒樟树港辣椒",
            "price": 48.0
          },
          {
            "name": "回锅牛肉",
            "price": 48.0
          },
          {
            "name": "四季豆炒牛肚",
            "price": 48.0
          },
          {
            "name": "小炒黄牛肉",
            "price": 48.0
          }
        ]
      },
      {
        "id": "p904",
        "name": "猪肉类-炒菜",
        "price": 28.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/BkMC6mGfir4iJ8ZZdcTNKpm5XjK43kHy.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不含主食"
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
        ],
        "specs": [
          {
            "name": "香干回锅肉",
            "price": 32.0
          },
          {
            "name": "樟树港辣椒炒肉",
            "price": 38.0
          },
          {
            "name": "青椒肉丝",
            "price": 28.0
          },
          {
            "name": "洋葱肉丝",
            "price": 28.0
          },
          {
            "name": "榨菜肉丝",
            "price": 28.0
          },
          {
            "name": "鱼香肉丝",
            "price": 38.0
          },
          {
            "name": "爆炒猪肝",
            "price": 28.0
          },
          {
            "name": "回锅肉炒卤藕",
            "price": 32.0
          },
          {
            "name": "爆炒顺风",
            "price": 38.0
          },
          {
            "name": "红烧蹄膀",
            "price": 38.0
          },
          {
            "name": "农家一碗香",
            "price": 32.0
          }
        ]
      },
      {
        "id": "p905",
        "name": "时蔬类-盖码份量",
        "price": 9.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/CYzYpFxrcTtbWk8zGAdKzTGmxAFNG2pk.jpg",
        "opts": [
          {
            "g": "现炒码子主食",
            "req": true,
            "opts": [
              "不含主食"
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
        ],
        "specs": [
          {
            "name": "爆炒阳雀湖甘蓝（农博会金奖）",
            "price": 12.0
          },
          {
            "name": "素炒小白菜",
            "price": 9.0
          },
          {
            "name": "清炒菜苔",
            "price": 12.0
          },
          {
            "name": "酸辣菜苔",
            "price": 12.0
          },
          {
            "name": "清炒泥蒿",
            "price": 16.0
          }
        ]
      }
    ]
  },
  {
    "id": "cat10",
    "name": "鲜货炭火烧烤",
    "items": [
      {
        "id": "p1001",
        "name": "碳烤精品羊肋条",
        "price": 8.0,
        "desc": "精选羊排，肉香十足，焦嫩多汁。",
        "tags": [],
        "img": "https://s.koubei.com/28t3YGpT64R2BwhabpcNFMEYawEdFD2P.jpg",
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
        "name": "原味羊肉串",
        "price": 5.0,
        "desc": "羊肉鲜嫩，炭烤香气浓郁，默认微辣",
        "tags": [],
        "img": "https://s.koubei.com/cEzesb6xNSKBBaPPDnFddwdMB4PGjQkM.jpg",
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
        "name": "原味猪肉串",
        "price": 3.0,
        "desc": "鲜肉现穿，清淡原味，肉汁丰富。",
        "tags": [],
        "img": "https://s.koubei.com/NDMbdWrxx5eStZp3dd86z7Tjf8je3SY6.jpg",
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
        "name": "牛肋条",
        "price": 6.0,
        "desc": "鲜牛肉肉质细嫩带微脆筋，烤后多汁，孜然肉香浓郁",
        "tags": [],
        "img": "https://s.koubei.com/YAwtR82DRnMQDTjD5zMTi8JaskTHQ42E.jpg",
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
        "name": "猪脆骨",
        "price": 3.0,
        "desc": "脆嫩有嚼劲，烧烤风味十足，默认微辣",
        "tags": [],
        "img": "https://s.koubei.com/yFYCKtfkxJFiSeZJkDfzaAGWP7ZTGcAi.jpg",
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
        "name": "王中王火腿肠",
        "price": 5.0,
        "desc": "外焦里嫩，经典烧烤风味，默认微辣",
        "tags": [],
        "img": "https://s.koubei.com/bhxC77G7GMcQKSNXiG8Tr7E4GkasDrd2.jpg",
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
        "name": "烤凤爪",
        "price": 12.0,
        "desc": "先卤后烤，软糯Q弹",
        "tags": [],
        "img": "https://s.koubei.com/WcyseBzfztAYMM625tkzYNyTAk2DzDCw.jpg",
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
        "name": "西葫芦",
        "price": 4.0,
        "desc": "烤后清甜，软嫩多汁，清爽解腻",
        "tags": [],
        "img": "https://s.koubei.com/Xa7EJmiyTkmniWEaWKwnJhaA3Wr23trn.jpg",
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
        "name": "烤鸡胗",
        "price": 3.0,
        "desc": "炭火现烤鸡胗，外皮微焦内里脆弹，满满粗孜然香，越嚼越上头。",
        "tags": [],
        "img": "https://s.koubei.com/brXKRs4HAXwyipBkBbN6G6dX5H5nW8Ey.jpg",
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
        "img": "https://s.koubei.com/ZnW6eiC4Y4dhiaKYhMnr8AYh43xm6rwM.jpg",
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
        "name": "香蕉吐司",
        "price": 8.0,
        "desc": "外酥内软，老少皆宜，微甜",
        "tags": [
          "咸甜"
        ],
        "img": "https://s.koubei.com/fG7tmp54KtBC7si8Atp8XBH2nAkWRpYE.jpg",
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
    "name": "手工鲜制饮品",
    "items": [
      {
        "id": "p1101",
        "name": "红枣姜汁汤圆",
        "price": 6.0,
        "desc": "暖姜微辛，红糖温润，一口暖到心底。",
        "tags": [],
        "img": "https://s.koubei.com/BzGd7y7wbRZ2HsR2zydKTGnMEfENn8Bc.jpg"
      },
      {
        "id": "p1102",
        "name": "绿豆汤",
        "price": 5.0,
        "desc": "清凉解暑，甜度适中",
        "tags": [],
        "img": "https://s.koubei.com/pHFiEzctsS5JkhH4mYfQQbdjwnp6d3nH.jpg"
      },
      {
        "id": "p1103",
        "name": "芒果珍珠豆花",
        "price": 8.0,
        "desc": "豆花嫩滑、芒果香甜、珍珠Q弹，果香与焦糖交织，层次丰富。",
        "tags": [],
        "img": "https://s.koubei.com/JRTWPBMi2iFASb5kxA7SDtD45nWyMQeM.jpg"
      },
      {
        "id": "p1104",
        "name": "木薯糖水",
        "price": 6.0,
        "desc": "木薯软糯，糖水温润",
        "tags": [],
        "img": "https://s.koubei.com/GQ4N2PZpdRdQYt4DBGPD2B2mNws3a3eh.jpg"
      },
      {
        "id": "p1105",
        "name": "桂花蜜豆花",
        "price": 6.0,
        "desc": "豆花滑嫩，桂花蜜清甜",
        "tags": [],
        "img": "https://s.koubei.com/5mZAe7bxNmWkzNp8FCp27JM432HTfw4y.jpg"
      },
      {
        "id": "p1106",
        "name": "桂花米酒汤圆",
        "price": 6.0,
        "desc": "酒酿鲜醇，圆子软糯，桂香淡淡的，甜而不齁。",
        "tags": [],
        "img": "https://s.koubei.com/ErBKwBztmEtnFa4dDxBjtCfKh55K5zKC.jpg"
      },
      {
        "id": "p1107",
        "name": "桂圆莲子百合汤",
        "price": 12.0,
        "desc": "桂圆清甜、莲子粉糯、百合软润，慢炖成清亮甜汤，温润顺口。",
        "tags": [],
        "img": "https://s.koubei.com/yY384acAhzARCateNTncPK473AYhR8Dx.jpg"
      }
    ]
  },
  {
    "id": "cat12",
    "name": "湖南特色炖锅",
    "items": [
      {
        "id": "p1201",
        "name": "特色牛杂炖粉",
        "price": 68.0,
        "desc": "新鲜牛肉、牛肚、牛筋、牛肠、牛黄喉配秘制锅底，细火慢炖，辣滚香鲜，汤醇味厚；汤汁通过高温融入米粉，入口软烂适中，咸淡适宜。",
        "tags": [],
        "img": "https://s.koubei.com/EcWt3SsT6NzeZ3y2kbm2HSbM5XR6pDcm.png",
        "opts": [
          {
            "g": "辣度",
            "req": true,
            "opts": [
              "微辣",
              "中辣",
              "不要辣"
            ]
          },
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "圆粉",
              "宽粉",
              "手工大宽粉",
              "扁粉",
              "湖南宽粉"
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
        "id": "p1202",
        "name": "光头粉（只能搭配炖锅下单）",
        "price": 5.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/EFJcCpm3bc7ZCb7RKXFZPx37FfXaKbBD.jpg",
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
        "id": "p1203",
        "name": "时蔬拼盘（配炖锅下单）",
        "price": 18.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/6RHfYsZYiz4MhXG7Yp8YBtrA5fwQPSFd.jpg"
      }
    ]
  },
  {
    "id": "cat13",
    "name": "时令鲜汤",
    "items": [
      {
        "id": "p1301",
        "name": "财鱼皮蛋豆腐汤",
        "price": 28.0,
        "desc": "财鱼鲜嫩，皮蛋提香，汤清味鲜",
        "tags": [
          "咸鲜"
        ],
        "img": "https://s.koubei.com/Hn7WEsZ34xmAb74GwKnhQy3sjaExjcaC.jpg",
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
        "id": "p1302",
        "name": "番茄蛋汤",
        "price": 18.0,
        "desc": "酸甜开胃，经典家常汤品",
        "tags": [
          "咸鲜",
          "微酸"
        ],
        "img": "https://s.koubei.com/mbrH8EpYSedY2dXHDEctKs4347dEy5jT.jpg",
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
        "id": "p1303",
        "name": "番茄肉片蛋汤",
        "price": 24.0,
        "desc": "番茄酸甜，肉片蛋花鲜嫩",
        "tags": [
          "咸鲜",
          "微酸"
        ],
        "img": "https://s.koubei.com/PfjpkBXF7rXMaR4XbSajbiKmWAcm8f7m.jpg",
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
        "id": "p1304",
        "name": "排骨藕汤/单人份",
        "price": 12.0,
        "desc": "排骨软烂，莲藕软糯，汤鲜味浓",
        "tags": [
          "咸鲜",
          "黑胡椒"
        ],
        "img": "https://s.koubei.com/tbX4QsBpkNMwniPMhJHJRSJPSaNpktn4.jpg",
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
        "id": "p1305",
        "name": "标准份-排骨藕汤",
        "price": 22.0,
        "desc": "适合1~2人分享，汤鲜味浓，滋补暖身",
        "tags": [
          "咸鲜",
          "黑胡椒"
        ],
        "img": "https://s.koubei.com/TieHtJRn8NBSRbdBixbzGwDmEAsFAEjN.jpg",
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
        "id": "p1306",
        "name": "丝瓜汤系列",
        "price": 18.0,
        "desc": "丝瓜清甜软嫩，牛肉滑嫩鲜香，汤清味鲜，清爽不腻。",
        "tags": [
          "咸鲜"
        ],
        "img": "https://s.koubei.com/SGTYaBmtSbWh3bRRbFdDBMsXMnjZbnEb.jpg",
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
            "price": 24.0
          },
          {
            "name": "丝瓜蛋汤",
            "price": 18.0
          }
        ]
      },
      {
        "id": "p1307",
        "name": "平菇肉片汤",
        "price": 3.0,
        "desc": "平菇鲜嫩、肉片滑嫩，汤清味鲜，暖胃爽口。",
        "tags": [
          "不可续"
        ],
        "img": "https://s.koubei.com/7ZsF4EkJNjzMGxGY4wFenmtjspDdtPNS.jpg",
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
      }
    ]
  },
  {
    "id": "cat14",
    "name": "卤菜凉菜",
    "items": [
      {
        "id": "p1401",
        "name": "凉拌卤牛肉",
        "price": 24.0,
        "desc": "卤牛肉醇香，凉拌鲜辣爽口",
        "tags": [],
        "img": "https://s.koubei.com/EPrbtHXfNteENMS2DNib6jGcdi5eZGRz.jpg",
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
        "id": "p1402",
        "name": "凉拌卤顺风",
        "price": 18.0,
        "desc": "卤猪耳脆嫩，凉拌香辣爽口",
        "tags": [],
        "img": "https://s.koubei.com/ayzFbT8aBiDwRh3FErDhxcF7dDBJWQ8z.jpg",
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
        "id": "p1403",
        "name": "皮蛋拌豆腐",
        "price": 12.0,
        "desc": "豆腐滑嫩，皮蛋鲜香开胃",
        "tags": [],
        "img": "https://s.koubei.com/GmBE82zW3tyw5B2RabSdDTGQPj4BchNd.jpg",
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
        "id": "p1404",
        "name": "凉拌黄瓜",
        "price": 8.0,
        "desc": "黄瓜爽脆，蒜香酸辣开胃",
        "tags": [],
        "img": "https://s.koubei.com/7D4ZJaTyB2AT26Mjwdc3eJWbDCrK84fe.jpg",
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
        "id": "p1405",
        "name": "麻辣蛤蜊",
        "price": 12.0,
        "desc": "蛤蜊鲜嫩，鲜甜入味",
        "tags": [],
        "img": "https://s.koubei.com/Ww5diGxkAR7ncPDZyjzDWasmNtimyGtk.jpg",
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
        "id": "p1406",
        "name": "凉拌卤藕",
        "price": 12.0,
        "desc": "卤藕粉糯脆爽，凉拌鲜香",
        "tags": [],
        "img": "https://s.koubei.com/nf5sJzpC5PYJaiwch4FQDGWe5ZemjMTa.jpg",
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
        "id": "p1407",
        "name": "凉拌卤拼盘",
        "price": 28.0,
        "desc": "多款卤味拼盘，咸香丰富",
        "tags": [],
        "img": "https://s.koubei.com/kfxZApYbfbD57hkE625d8pGAcfEczSdp.jpg",
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
    "id": "cat15",
    "name": "炒粉炒面炒饭炒豆丝",
    "items": [
      {
        "id": "p1501",
        "name": "炒粉面饭类",
        "price": 18.0,
        "desc": "",
        "tags": [
          "做法原因不可续加"
        ],
        "img": "https://s.koubei.com/JM7W5sBTtbrMdM2pBT4RRnyttRHkDihP.jpg",
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
              "【打包】品牌餐盒",
              "其他需求请备注"
            ]
          }
        ],
        "specs": [
          {
            "name": "鲜猪肉",
            "price": 22.0
          },
          {
            "name": "鸡蛋火腿",
            "price": 18.0
          },
          {
            "name": "鲜牛肉",
            "price": 32.0
          },
          {
            "name": "卤牛肉",
            "price": 32.0
          },
          {
            "name": "卤顺风",
            "price": 24.0
          },
          {
            "name": "腊肠",
            "price": 24.0
          }
        ]
      }
    ]
  },
  {
    "id": "cat16",
    "name": "主食",
    "items": [
      {
        "id": "p1601",
        "name": "五常大米饭",
        "price": 3.0,
        "desc": "五常大米蒸制，米香软糯",
        "tags": [],
        "img": "https://s.koubei.com/QYcs7KEs2sJTWEMdbX2F3EWPkBzbXeGA.jpg"
      }
    ]
  },
  {
    "id": "cat17",
    "name": "酒水饮料",
    "items": [
      {
        "id": "p1701",
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
        "id": "p1702",
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
        "id": "p1703",
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
        "id": "p1704",
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
        "id": "p1705",
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
        "id": "p1706",
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
        "id": "p1707",
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
        "id": "p1708",
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
        "id": "p1709",
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
        "id": "p1710",
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
        "id": "p1711",
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
        "id": "p1712",
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
        "id": "p1713",
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
        "id": "p1714",
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
        "id": "p1715",
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
        "id": "p1716",
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
        "id": "p1717",
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
        "id": "p1718",
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
        "id": "p1719",
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
        "id": "p1720",
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
