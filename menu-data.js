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
 * 注意：以下分类ID与餐盒逻辑绑定，请勿修改：
 *   cat1  -> 牛肉粉面（干湿分离盒 2元）
 *   cat2  -> 鲜鱼汤（码子干湿分离 2.5元）
 *   cat5  -> 虾蟹面（品牌餐盒 1元）
 *   cat6  -> 家常炒菜（品牌餐盒 1元）
 *   cat7  -> 新品专区（品牌餐盒 1元）
 *   cat10 -> 卤菜凉菜（码子方盒 0.5元）
 *   cat11 -> 时令鲜汤（品牌餐盒 1元）
 *   cat12 -> 盖码饭（品牌餐盒 1元）
 *   cat13 -> 炒粉炒面炒豆丝（品牌餐盒 1元）
 */

const MENU = [
  {
    "id": "cat1",
    "name": "招牌牛肉粉面",
    "items": [
      {
        "id": "p101",
        "name": "招牌红烧牛肉粉",
        "price": 25.0,
        "desc": "慢炖坨状牛腩，软糯Q弹，老少皆宜",
        "tags": [
          "咪咪辣"
        ],
        "img": "https://s.koubei.com/BjHyGWcXMk3QCii2TYy7zRa78iHAahQf.jpg",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "圆粉",
              "扁粉",
              "宽粉",
              "手工大宽粉",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "碱面",
              "豆丝"
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
              "加-卤鸡爪/只",
              "加-卤鸭掌/只"
            ]
          },
          {
            "g": "粉面（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】干湿分离打包盒",
              "【外带】品牌餐盒"
            ]
          }
        ]
      },
      {
        "id": "p102",
        "name": "秘制牛肉粉",
        "price": 25.0,
        "desc": "文火慢卤切片，汤底醇厚，微辣鲜香",
        "tags": [
          "微辣",
          "招牌必点"
        ],
        "img": "https://s.koubei.com/s8efjaBMyH2eKB3WicSYQ43CkPJPnRrA.jpg",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "圆粉",
              "扁粉",
              "宽粉",
              "手工大宽粉",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "碱面",
              "豆丝"
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
              "加-卤鸡爪/只",
              "加-卤鸭掌/只"
            ]
          },
          {
            "g": "粉面（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】干湿分离打包盒",
              "【外带】品牌餐盒"
            ]
          }
        ]
      },
      {
        "id": "p103",
        "name": "特色牛三鲜粉",
        "price": 25.0,
        "desc": "脆爽与软糯交织，中辣过瘾",
        "tags": [
          "中辣"
        ],
        "img": "https://s.koubei.com/aswSnS7dkPwFbsYY4PCZrEZXdQmTyJMX.jpg",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "圆粉",
              "扁粉",
              "宽粉",
              "手工大宽粉",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "碱面",
              "豆丝"
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
              "加-卤鸡爪/只",
              "加-卤鸭掌/只"
            ]
          },
          {
            "g": "粉面（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】干湿分离打包盒",
              "【外带】品牌餐盒"
            ]
          }
        ]
      },
      {
        "id": "p104",
        "name": "热干面系列",
        "price": 8.0,
        "desc": "默认微辣",
        "tags": [
          "做法原因不可续面"
        ],
        "img": "https://s.koubei.com/A4kGxje6y3ByYGdFJwdjCzEb2AyzNaPh.jpg",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "热干面",
              "热干-宽粉",
              "热干-扁粉",
              "热干-圆粉"
            ]
          },
          {
            "g": "加码小吃",
            "req": false,
            "opts": [
              "加-卤蛋/枚",
              "加-海带结2个",
              "加-卤鸡爪/只",
              "加-卤鸭掌/只"
            ]
          },
          {
            "g": "粉面（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】品牌餐盒"
            ]
          }
        ]
      }
    ]
  },
  {
    "id": "cat2",
    "name": "鲜鱼汤-现炒码子粉",
    "items": [
      {
        "id": "p201",
        "name": "鲜鱼汤-财鱼粉",
        "price": 32.0,
        "desc": "新鲜财鱼肉质紧实脆嫩，配鲜鱼汤底，鲜甜微辣",
        "tags": [],
        "img": "https://s.koubei.com/GEYwjNeZsyd3eP35FG2FHcGepzBEr66F.jpg",
        "opts": [
          {
            "g": "码子粉主食",
            "req": true,
            "opts": [
              "圆粉",
              "扁粉",
              "宽粉",
              "手工大宽粉",
              "碱面",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "豆丝"
            ]
          },
          {
            "g": "码子粉（默认堂食不用选）",
            "req": true,
            "opts": [
              "【外带】牛皮纸碗",
              "默认堂食碗",
              "【外带】码子粉干湿分离",
              "【外带】品牌餐盒"
            ]
          }
        ]
      },
      {
        "id": "p202",
        "name": "鲜鱼汤-鳝鱼粉",
        "price": 48.0,
        "desc": "口感脆嫩，配鲜鱼汤底，来一碗舒服！",
        "tags": [],
        "img": "https://s.koubei.com/GmcKmPrYGzEwhPzxZK2p3wpEMMdNB3Ce.jpg",
        "opts": [
          {
            "g": "码子粉主食",
            "req": true,
            "opts": [
              "圆粉",
              "扁粉",
              "宽粉",
              "手工大宽粉",
              "碱面",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "豆丝"
            ]
          },
          {
            "g": "码子粉（默认堂食不用选）",
            "req": true,
            "opts": [
              "【外带】牛皮纸碗",
              "默认堂食碗",
              "【外带】码子粉干湿分离",
              "【外带】品牌餐盒"
            ]
          }
        ]
      },
      {
        "id": "p203",
        "name": "鲜鱼汤-鲜牛肉粉",
        "price": 36.0,
        "desc": "鲜牛肉与鱼汤融合，肉香与鲜甜交织",
        "tags": [],
        "img": "https://s.koubei.com/6iktsC3s7sZxw6MTfAipnf22KkwCsGa2.jpg",
        "opts": [
          {
            "g": "码子粉主食",
            "req": true,
            "opts": [
              "圆粉",
              "扁粉",
              "宽粉",
              "手工大宽粉",
              "碱面",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "豆丝"
            ]
          },
          {
            "g": "鲜鱼汤码子粉",
            "req": true,
            "opts": [
              "新做法:ccm:不可免辣"
            ]
          },
          {
            "g": "码子粉（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】品牌餐盒",
              "【外带】码子粉干湿分离"
            ]
          }
        ]
      },
      {
        "id": "p204",
        "name": "鲜鱼汤-瘦肉粉",
        "price": 22.0,
        "desc": "鲜猪肉配薄皮青椒，锅气十足，鲜甜微辣",
        "tags": [],
        "img": "https://s.koubei.com/CfYeFafw5TnixE2WnmRHcQpPynW8y4Kc.png",
        "opts": [
          {
            "g": "码子粉主食",
            "req": true,
            "opts": [
              "圆粉",
              "扁粉",
              "宽粉",
              "手工大宽粉",
              "碱面",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "豆丝"
            ]
          },
          {
            "g": "码子粉（默认堂食不用选）",
            "req": true,
            "opts": [
              "【外带】牛皮纸碗",
              "默认堂食碗",
              "【外带】码子粉干湿分离",
              "【外带】品牌餐盒"
            ]
          }
        ]
      },
      {
        "id": "p205",
        "name": "鲜鱼汤-猪肝粉",
        "price": 20.0,
        "desc": "猪肝嫩滑，配青椒洋葱，汁浓味美",
        "tags": [],
        "img": "https://s.koubei.com/6MhNaNzWisnJ6fZtby7TrKb8Rjtswytz.jpg",
        "opts": [
          {
            "g": "码子粉主食",
            "req": true,
            "opts": [
              "圆粉",
              "扁粉",
              "宽粉",
              "手工大宽粉",
              "碱面",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "豆丝"
            ]
          },
          {
            "g": "码子粉（默认堂食不用选）",
            "req": true,
            "opts": [
              "【外带】牛皮纸碗",
              "默认堂食碗",
              "【外带】码子粉干湿分离",
              "【外带】品牌餐盒"
            ]
          }
        ]
      },
      {
        "id": "p206",
        "name": "鲜鱼汤-瘦肉猪肝粉",
        "price": 24.0,
        "desc": "瘦肉与猪肝双拼，鲜嫩加倍，汤底鲜甜",
        "tags": [],
        "img": "https://s.koubei.com/jQ7CrN4E5xmNmTNwzCQsNSFz44wyzayz.jpg",
        "opts": [
          {
            "g": "码子粉主食",
            "req": true,
            "opts": [
              "圆粉",
              "扁粉",
              "宽粉",
              "手工大宽粉",
              "碱面",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "豆丝"
            ]
          },
          {
            "g": "码子粉（默认堂食不用选）",
            "req": true,
            "opts": [
              "【外带】牛皮纸碗",
              "默认堂食碗",
              "【外带】码子粉干湿分离",
              "【外带】品牌餐盒"
            ]
          }
        ]
      },
      {
        "id": "p207",
        "name": "金汤黄骨鱼粉",
        "price": 36.0,
        "desc": "制作工艺原因，出餐时间需20分钟以上，如果时间着急，不妨下次再来品尝这道菜",
        "tags": [
          "新品"
        ],
        "img": "https://s.koubei.com/HSiCCj8Q27d36cHNGHNdPtjdkbZXRmPA.jpg",
        "opts": [
          {
            "g": "码子粉主食",
            "req": true,
            "opts": [
              "圆粉",
              "扁粉",
              "宽粉",
              "手工大宽粉",
              "碱面",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "豆丝"
            ]
          },
          {
            "g": "出餐说明",
            "req": true,
            "opts": [
              "制作工艺原因:ccm:出餐时长20-30分钟"
            ]
          },
          {
            "g": "码子粉（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】1768打包盒"
            ]
          }
        ]
      },
      {
        "id": "p208",
        "name": "鲜鱼汤-鲜虾粉",
        "price": 36.0,
        "desc": "鲜活基围虾，肉紧弹牙，汤底鲜甜，微辣提味",
        "tags": [],
        "img": "https://s.koubei.com/HrCb4MY4hZzXS6FT2dHdGZHznYwJRFd5.jpg",
        "opts": [
          {
            "g": "码子粉主食",
            "req": true,
            "opts": [
              "圆粉",
              "扁粉",
              "宽粉",
              "手工大宽粉",
              "碱面",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "豆丝"
            ]
          },
          {
            "g": "鲜鱼汤码子粉",
            "req": true,
            "opts": [
              "新做法:ccm:不可免辣"
            ]
          },
          {
            "g": "码子粉（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】品牌餐盒",
              "【外带】码子粉干湿分离"
            ]
          }
        ]
      }
    ]
  },
  {
    "id": "cat3",
    "name": "特色小吃",
    "items": [
      {
        "id": "p301",
        "name": "卤海带结2个",
        "price": 3.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/rTr5xDyFGy7yJN7NeM7ctWpArRjf5xQc.jpg"
      },
      {
        "id": "p302",
        "name": "牛肉汤卤蛋/枚",
        "price": 3.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/eBem3bKsBhEPjrCasNeRRiGN5eBj6Hi6.jpg"
      },
      {
        "id": "p303",
        "name": "卤鸡爪/只",
        "price": 6.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/jjMxhxp8GDjCJSiRiPJSWR3xWH2JSfMB.jpg"
      },
      {
        "id": "p304",
        "name": "卤鸭掌/只",
        "price": 6.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/5xY7E2xneYTwpyxEswejeeycEDM8dDfK.jpg"
      }
    ]
  },
  {
    "id": "cat4",
    "name": "手工鲜制饮品",
    "items": [
      {
        "id": "p401",
        "name": "雪顶青茉",
        "price": 15.0,
        "desc": "口味绵密如慕斯，浓郁奶香，抹茶微苦回甘",
        "tags": [],
        "img": "https://s.koubei.com/repDCGPBCM7XKASRQy83Y8W8ytb5iMfa.jpg"
      },
      {
        "id": "p402",
        "name": "琥珀冬瓜茶",
        "price": 12.0,
        "desc": "冬瓜清香，甜而不腻",
        "tags": [],
        "img": "https://s.koubei.com/t5cDAMG2e6fNnZBm2tKwnTr7eXbhTEw2.jpg"
      },
      {
        "id": "p403",
        "name": "桂花蜜豆花",
        "price": 6.0,
        "desc": "豆花滑嫩，桂花蜜清甜",
        "tags": [],
        "img": "https://s.koubei.com/SH7NEyhzH6W4w3XXD6RZnHSjm7z7HiBi.jpg"
      },
      {
        "id": "p404",
        "name": "绿豆汤",
        "price": 5.0,
        "desc": "清凉解暑，甜度适中",
        "tags": [],
        "img": "https://s.koubei.com/zZRmsCsFy3MA3DZSG7RF6d3SptxcY4Gx.jpg"
      },
      {
        "id": "p405",
        "name": "木薯糖水",
        "price": 6.0,
        "desc": "木薯软糯，糖水温润",
        "tags": [],
        "img": "https://s.koubei.com/EfXPS2mm8EaMnARbSpXnjC2ExFwzYaM5.jpg",
        "opts": [
          {
            "g": "温度",
            "req": true,
            "opts": [
              "冰",
              "常温"
            ]
          }
        ]
      },
      {
        "id": "p406",
        "name": "鸭屎香蜂蜜双柠茶",
        "price": 12.0,
        "desc": "鸭屎香单丛茶底兰香清雅，双柠果香清爽，蜂蜜甜润自然，三香交融顺滑解腻。",
        "tags": [],
        "img": "https://s.koubei.com/hBTaJejZQGGJZsrNYncBTkXTcS6eYQi5.jpg"
      }
    ]
  },
  {
    "id": "cat5",
    "name": "新品-蟹钳面",
    "items": [
      {
        "id": "p501",
        "name": "秘制蟹钳面",
        "price": 48.0,
        "desc": "武汉一绝！入口先甜、中后辣、尾调带麻，鲜、香、麻、辣层层递进",
        "tags": [
          "推荐",
          "做法原因不可续面"
        ],
        "img": "https://s.koubei.com/hk2pBSZwxcrX73eKGtD8M6E5tPCjhXsC.jpg"
      },
      {
        "id": "p502",
        "name": "炒菜-秘制蟹钳",
        "price": 78.0,
        "desc": "武汉一绝！入口先甜、中后辣、尾调带麻，鲜、香、麻、辣层层递进",
        "tags": [],
        "img": "https://s.koubei.com/xTXc3rmNch23dYekDs5XHDE3DJpNPjhi.jpg"
      }
    ]
  },
  {
    "id": "cat6",
    "name": "家常炒菜",
    "items": [
      {
        "id": "p601",
        "name": "炒菜-鲜牛肉炒辣椒",
        "price": 48.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/zBTXAGEwNawc7b2Zanp2MZdbQAzHW8h6.jpg"
      },
      {
        "id": "p602",
        "name": "炒菜-农家炒牛肉丝",
        "price": 48.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/c8sCfpK7nAKzjNjTQwQDcryc3kfT7aJy.jpg"
      },
      {
        "id": "p603",
        "name": "炒菜-仔姜牛肉丝",
        "price": 48.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/h5368CFiRaws3xMXCFypbHZNFPFpwkRt.jpg"
      },
      {
        "id": "p604",
        "name": "炒菜-酸萝卜炒牛肉",
        "price": 48.0,
        "desc": "酸萝卜酸甜脆爽，牛肉鲜嫩，酸辣开胃",
        "tags": [],
        "img": "https://s.koubei.com/St6Y4jzk5j7NMCZbyXTrT6ARwtmwBRMN.jpg"
      },
      {
        "id": "p605",
        "name": "炒菜-回锅牛肉",
        "price": 48.0,
        "desc": "卤牛肉干香醇厚，软而不柴，酱香浓郁",
        "tags": [],
        "img": "https://s.koubei.com/S4FhmmrQ5wie6NxNHtycKeS676JpefeF.jpg"
      },
      {
        "id": "p606",
        "name": "炒菜-卤藕炒回锅牛肉",
        "price": 48.0,
        "desc": "卤肉卤藕，粉糯干香，微辣醇厚",
        "tags": [],
        "img": "https://s.koubei.com/Rnz63zYWEGHh2CBw323s2NSFhsGeijPY.jpg"
      },
      {
        "id": "p607",
        "name": "炒菜-四季豆炒牛肚",
        "price": 38.0,
        "desc": "牛肚脆爽，四季豆鲜嫩，中辣咸鲜",
        "tags": [],
        "img": "https://s.koubei.com/Ajxzh27cM3b5zE7BMwQJYeeFXMz3hzRe.jpg"
      },
      {
        "id": "p608",
        "name": "炒菜-青椒焖牛腩",
        "price": 48.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/zFQC2GhKXTRiYBWQpfGptX3x2WFS8yKW.jpg"
      },
      {
        "id": "p609",
        "name": "炒菜-樟树港辣椒炒肉",
        "price": 38.0,
        "desc": "樟树港辣椒别具清香，肉片鲜嫩，微辣下饭",
        "tags": [],
        "img": "https://s.koubei.com/zTiTAtZwCSPtMCn76sRsSjjQyMbS4PCp.jpg"
      },
      {
        "id": "p610",
        "name": "炒菜-香干回锅肉",
        "price": 32.0,
        "desc": "五花肉肥而不腻，香干外韧里嫩，酱香微辣",
        "tags": [],
        "img": "https://s.koubei.com/piyMHzZbeznKsipJDYTExm8Bi5Ksp6az.jpg"
      },
      {
        "id": "p611",
        "name": "炒菜-回锅肉炒卤藕",
        "price": 32.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/hNZK4zFJMe2zWnsrJ2hCXG8hr7ithZ3R.jpg"
      },
      {
        "id": "p612",
        "name": "炒菜-爆炒顺风",
        "price": 38.0,
        "desc": "卤猪耳脆嫩爽口，咸鲜香辣，锅气十足",
        "tags": [],
        "img": "https://s.koubei.com/eHM4eKD7ZTJMQYZrb6Fn8jZpFBsSdjGs.jpg"
      },
      {
        "id": "p613",
        "name": "炒菜-青椒肉丝",
        "price": 28.0,
        "desc": "青椒脆嫩，肉丝滑嫩，咸鲜微辣",
        "tags": [],
        "img": "https://s.koubei.com/jP8eyf63jDHFtTf5CMx2PmMzpGjcfXAc.jpg"
      },
      {
        "id": "p614",
        "name": "炒菜-洋葱肉丝",
        "price": 28.0,
        "desc": "洋葱清甜，肉丝嫩滑，咸鲜适口",
        "tags": [],
        "img": "https://s.koubei.com/f4bJb6MKWKMsp5YzzinMcx3ecDjahbHD.jpg"
      },
      {
        "id": "p615",
        "name": "炒菜-榨菜肉丝",
        "price": 28.0,
        "desc": "双椒鲜辣，榨菜咸香，下饭神器",
        "tags": [],
        "img": "https://s.koubei.com/mYkRpJ5zrfsGbTdyx5TYk2N2T6axw677.jpg"
      },
      {
        "id": "p616",
        "name": "炒菜-鱼香肉丝",
        "price": 38.0,
        "desc": "肉丝滑嫩，，酸甜辣兼具",
        "tags": [],
        "img": "https://s.koubei.com/JxMWRi8XGaMNNRPFJ52K2SmQtxfbJAGJ.jpg"
      },
      {
        "id": "p617",
        "name": "炒菜-丝瓜炒肉片",
        "price": 32.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/jRc3dJfZAGAT7XjQNCsBKBB3ckb7DdWA.jpg"
      },
      {
        "id": "p618",
        "name": "炒菜-爆炒猪肝",
        "price": 28.0,
        "desc": "猪肝嫩滑，火候精准，咸鲜适口",
        "tags": [],
        "img": "https://s.koubei.com/QANEib3p4i4XnyjH2PEww6JYBJAwFdNJ.jpg"
      },
      {
        "id": "p619",
        "name": "炒菜-爆炒鳝鱼丝",
        "price": 88.0,
        "desc": "鳝鱼丝滑嫩，咸鲜微辣，风味浓郁",
        "tags": [],
        "img": "https://s.koubei.com/GEjNy6KZEyKCy8pK5RixszERWzyaPnsG.jpg"
      },
      {
        "id": "p620",
        "name": "炒菜-爆炒财鱼片",
        "price": 58.0,
        "desc": "财鱼片脆嫩紧实，汤汁浓郁，鲜甜入味",
        "tags": [],
        "img": "https://s.koubei.com/4dT56CTyM7HN6FWPXX5dzCGTsC2skMPk.jpg"
      },
      {
        "id": "p621",
        "name": "炒菜-番茄炒蛋",
        "price": 24.0,
        "desc": "酸甜开胃，经典家常味",
        "tags": [],
        "img": "https://s.koubei.com/7BNNiPn2NiFGiSjXbXk7CzPr5Jx5CKWk.jpg"
      },
      {
        "id": "p622",
        "name": "炒菜-火爆鸡胗",
        "price": 32.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/dhrcDseyyezSWP6YTxeWpZmipiJ8tRhz.jpg"
      },
      {
        "id": "p623",
        "name": "炒菜-青椒焖鸡块",
        "price": 38.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/iMceeW3GTPNa3mBXmQiiyFKZFheAbkta.jpg"
      },
      {
        "id": "p624",
        "name": "炒菜-爆炒鲜牛肉",
        "price": 48.0,
        "desc": "鲜牛肉猛火爆炒，滑嫩香辣，锅气十足",
        "tags": [],
        "img": "https://s.koubei.com/myQt6xeAzRMxmErRMERxcy5tkdtTcn2B.jpg"
      },
      {
        "id": "p625",
        "name": "炒菜-小炒黄牛肉",
        "price": 48.0,
        "desc": "鲜牛肉猛火爆炒，滑嫩鲜辣，油亮诱人。",
        "tags": [],
        "img": "https://s.koubei.com/yjksP7TPFeaypxpzWPf2TybD5TsnPTb5.jpg"
      }
    ]
  },
  {
    "id": "cat7",
    "name": "新品专区",
    "items": [
      {
        "id": "p701",
        "name": "炒菜-青椒茄子",
        "price": 22.0,
        "desc": "茄子软糯，青椒清香",
        "tags": [],
        "img": "https://s.koubei.com/zkJMYMASxKeAEwmY8a5dZZz5nmiBJ7t3.jpg"
      },
      {
        "id": "p702",
        "name": "炒菜-清炒油麦菜",
        "price": 18.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/ZidtRDpDBenxD3JiBTRbcN2Qj3YynR5D.jpg"
      },
      {
        "id": "p703",
        "name": "炒菜-酸辣藕带",
        "price": 26.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/AFD6PyFkSNhcW8W5G2CsKrY8dnWbZ7kw.jpg"
      },
      {
        "id": "p704",
        "name": "炒菜-清炒藕带",
        "price": 26.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/Fp5knniy2bra7JSeiSf2rFs8Z5pYph4A.jpg"
      },
      {
        "id": "p705",
        "name": "炒菜-清炒红苋菜",
        "price": 18.0,
        "desc": "口感软嫩鲜香，自带草本清甜风味，蒜香提味，清爽不腻口。",
        "tags": [],
        "img": "https://s.koubei.com/pd2SW2NF4n6AykmPaF2YTA8ym6iip3wy.jpg"
      },
      {
        "id": "p706",
        "name": "炒菜-清炒苕尖",
        "price": 18.0,
        "desc": "口感脆嫩爽口，自带清甜本味。蒜香鲜香入味，整体清淡解腻。",
        "tags": [],
        "img": "https://s.koubei.com/fhYDEpQxHsRT6DrXQTbGb3Zh8kfNW2H6.jpg"
      },
      {
        "id": "p707",
        "name": "炒菜-清炒空心菜",
        "price": 18.0,
        "desc": "口感脆嫩水润，自带天然清鲜，蒜香相融提味，清爽适口不厚重。",
        "tags": [],
        "img": "https://s.koubei.com/F2FPeMEr6StBFHRpGc2mkY2iKHYw4NKW.jpg"
      },
      {
        "id": "p708",
        "name": "炒菜-瓠子烧五花肉",
        "price": 22.0,
        "desc": "肉香打底，蒜香突出，瓠子自带清甜",
        "tags": [],
        "img": "https://s.koubei.com/EKK4KdGiKD86Ef37F4Shhw3HxbTkCTAs.jpg"
      },
      {
        "id": "p709",
        "name": "炒菜-家常炒丝瓜",
        "price": 22.0,
        "desc": "口味清淡，辣度柔和，老少皆宜",
        "tags": [],
        "img": "https://s.koubei.com/HtrhyEQznDbeFcHJYhRDR8t54A3pxDxz.jpg"
      },
      {
        "id": "p710",
        "name": "炒菜-擂椒烧豆角",
        "price": 22.0,
        "desc": "豆角脆韧有嚼劲，擂椒软糯焦香，软硬层次丰富",
        "tags": [],
        "img": "https://s.koubei.com/6tp4iNTBjkX57XQWiTHccQbGpacXzhei.jpg"
      }
    ]
  },
  {
    "id": "cat8",
    "name": "鲜货炭火烧烤",
    "items": [
      {
        "id": "p801",
        "name": "牛肉串",
        "price": 5.0,
        "desc": "鲜切牛肉，碳烤焦香，默认微辣",
        "tags": [],
        "img": "https://s.koubei.com/8QRStMKRecZsSQbHTFcR3tWxAcMxXQFA.jpg",
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
            "g": "烧烤做法",
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
    "id": "cat9",
    "name": "湖鲜牛杂炖锅",
    "items": [
      {
        "id": "p901",
        "name": "湖鲜牛杂炖锅",
        "price": 98.0,
        "desc": "鲜活基围虾、鲜切财鱼片、鲜卤牛杂，汤底浓郁，适合多人分享",
        "tags": [],
        "img": "https://s.koubei.com/EPrkH8FbzHMyeGTnF4caDXxw7zk5yNMy.jpg",
        "opts": [
          {
            "g": "码子粉主食",
            "req": true,
            "opts": [
              "圆粉",
              "扁粉",
              "宽粉",
              "手工大宽粉"
            ]
          },
          {
            "g": "辣度",
            "req": true,
            "opts": [
              "微辣",
              "中辣",
              "特辣"
            ]
          },
          {
            "g": "炖锅（默认堂食不用选）",
            "req": true,
            "opts": [
              "【外带】炖锅打包盒",
              "默认堂食碗"
            ]
          }
        ]
      }
    ]
  },
  {
    "id": "cat10",
    "name": "卤菜凉菜",
    "items": [
      {
        "id": "p1001",
        "name": "凉拌卤牛肉",
        "price": 24.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/Rnbc8HMn5ZHyaXTC8Ad4ZiHDaN7KwmaY.jpg",
        "opts": [
          {
            "g": "小吃例汤（默认堂食不用选）",
            "req": true,
            "opts": [
              "【外带】码子方盒",
              "默认堂食碗"
            ]
          }
        ]
      },
      {
        "id": "p1002",
        "name": "凉拌卤顺风",
        "price": 18.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/pjbbPcbsCx7D4tDcwD5e3WkWXsMYPmnN.jpg",
        "opts": [
          {
            "g": "小吃例汤（默认堂食不用选）",
            "req": true,
            "opts": [
              "【外带】码子方盒",
              "默认堂食碗"
            ]
          }
        ]
      },
      {
        "id": "p1003",
        "name": "皮蛋拌豆腐",
        "price": 12.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/7FCZmMxPTDwCTf3aBfDR64thbpm28dmr.png",
        "opts": [
          {
            "g": "小吃例汤（默认堂食不用选）",
            "req": true,
            "opts": [
              "【外带】码子方盒",
              "默认堂食碗"
            ]
          }
        ]
      },
      {
        "id": "p1004",
        "name": "凉拌黄瓜",
        "price": 8.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/JzY2zcGw3WJ6k7TZsfxdhS6jPaiR5NtJ.jpg",
        "opts": [
          {
            "g": "小吃例汤（默认堂食不用选）",
            "req": true,
            "opts": [
              "【外带】码子方盒",
              "默认堂食碗"
            ]
          }
        ]
      },
      {
        "id": "p1005",
        "name": "麻辣蛤蜊",
        "price": 12.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/Qmx5xcGrMZbXBpXRXNH72isjhckFHDHF.jpg",
        "opts": [
          {
            "g": "小吃例汤（默认堂食不用选）",
            "req": true,
            "opts": [
              "【外带】码子方盒",
              "默认堂食碗"
            ]
          }
        ]
      },
      {
        "id": "p1006",
        "name": "凉拌卤藕",
        "price": 12.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/jbJJ3tk8KbJRZbRRJGAD7Et2PhZCzArM.jpg",
        "opts": [
          {
            "g": "小吃例汤（默认堂食不用选）",
            "req": true,
            "opts": [
              "【外带】码子方盒",
              "默认堂食碗"
            ]
          }
        ]
      },
      {
        "id": "p1007",
        "name": "凉拌毛豆",
        "price": 12.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/fWmdBN5r7cfG47m74DTkSYxaFMaeAxPH.jpg",
        "opts": [
          {
            "g": "小吃例汤（默认堂食不用选）",
            "req": true,
            "opts": [
              "【外带】码子方盒",
              "默认堂食碗"
            ]
          }
        ]
      },
      {
        "id": "p1008",
        "name": "凉拌卤拼盘",
        "price": 28.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/Wd582KQJxKRX5ewCfypPHbXx38aAyhZZ.jpg",
        "opts": [
          {
            "g": "小吃例汤（默认堂食不用选）",
            "req": true,
            "opts": [
              "【外带】码子方盒",
              "默认堂食碗"
            ]
          }
        ]
      },
      {
        "id": "p1009",
        "name": "养生黑木耳",
        "price": 8.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/BP5ZPQKzMexScAne88ze5hGrT8CTziEj.jpg",
        "opts": [
          {
            "g": "小吃例汤（默认堂食不用选）",
            "req": true,
            "opts": [
              "【外带】码子方盒",
              "默认堂食碗"
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
        "tags": [],
        "img": "https://s.koubei.com/4bhXSsdGzZd5AYJCmimKGCXMEDCZm6sP.jpg",
        "opts": [
          {
            "g": "小吃例汤（默认堂食不用选）",
            "req": true,
            "opts": [
              "【外带】品牌餐盒",
              "默认堂食碗"
            ]
          }
        ]
      },
      {
        "id": "p1102",
        "name": "番茄蛋汤",
        "price": 18.0,
        "desc": "酸甜开胃，经典家常汤品",
        "tags": [],
        "img": "https://s.koubei.com/EZZ6aJ2H2kWFGxwW5WnTNArcHdJcn7kp.jpg",
        "opts": [
          {
            "g": "小吃例汤（默认堂食不用选）",
            "req": true,
            "opts": [
              "【外带】品牌餐盒",
              "默认堂食碗"
            ]
          }
        ]
      },
      {
        "id": "p1103",
        "name": "玉米排骨汤/单人份",
        "price": 12.0,
        "desc": "玉米清甜，排骨软烂，汤鲜味浓",
        "tags": [],
        "img": "https://s.koubei.com/h4EHs3CySD5YCXsAb6zA4pwJG8ZbjENK.jpg",
        "opts": [
          {
            "g": "小吃例汤（默认堂食不用选）",
            "req": true,
            "opts": [
              "【外带】品牌餐盒（小）",
              "默认堂食碗"
            ]
          }
        ]
      },
      {
        "id": "p1104",
        "name": "牛腩萝卜汤/单人份",
        "price": 16.0,
        "desc": "牛腩软烂，萝卜清甜，汤醇味美",
        "tags": [],
        "img": "https://s.koubei.com/StweXbfiW2GkrzDfhM4rYrsbiYRjdx5K.jpg",
        "opts": [
          {
            "g": "小吃例汤（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】品牌餐盒（小）"
            ]
          }
        ]
      },
      {
        "id": "p1105",
        "name": "标准份-玉米排骨汤",
        "price": 28.0,
        "desc": "适合2~3人分享，汤鲜味浓，滋补暖身",
        "tags": [],
        "img": "https://s.koubei.com/F3B8NfXWkYi5xte7byFjsDP3YpR5JcNi.jpg",
        "opts": [
          {
            "g": "小吃例汤（默认堂食不用选）",
            "req": true,
            "opts": [
              "【外带】品牌餐盒",
              "默认堂食碗"
            ]
          }
        ]
      },
      {
        "id": "p1106",
        "name": "标准份-牛腩萝卜汤",
        "price": 38.0,
        "desc": "适合2~3人分享，牛腩软烂，萝卜清甜，汤醇味美",
        "tags": [],
        "img": "https://s.koubei.com/3PWGkrGyz8ZWbfyzhRp3KpPeMRcSfRmk.jpg",
        "opts": [
          {
            "g": "小吃例汤（默认堂食不用选）",
            "req": true,
            "opts": [
              "【外带】品牌餐盒",
              "默认堂食碗"
            ]
          }
        ]
      },
      {
        "id": "p1107",
        "name": "番茄肉片蛋汤",
        "price": 24.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/rBRK8ZkY5QKRxzKSDKfJetCJt5nFhbQF.jpg",
        "opts": [
          {
            "g": "小吃例汤（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】品牌餐盒"
            ]
          }
        ]
      }
    ]
  },
  {
    "id": "cat12",
    "name": "盖码饭（饭菜一起）",
    "items": [
      {
        "id": "p1201",
        "name": "盖码饭-鲜牛肉炒辣椒",
        "price": 32.0,
        "desc": "辣椒鲜香，牛肉片嫩，微辣下饭",
        "tags": [],
        "img": "https://s.koubei.com/emQb3c2E6t8Dznt2MYwGm5cBTWTHDDJY.jpg",
        "opts": [
          {
            "g": "现炒码子（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】品牌餐盒",
              "【外带】牛皮纸碗分离",
              "【外带】干湿分离打包盒"
            ]
          }
        ]
      },
      {
        "id": "p1202",
        "name": "盖码饭-农家炒牛肉丝",
        "price": 32.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/mnPT87f3yxEhJNaXDkCb4i7fBNTi62hx.jpg",
        "opts": [
          {
            "g": "现炒码子（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】品牌餐盒",
              "【外带】牛皮纸碗分离",
              "【外带】干湿分离打包盒"
            ]
          }
        ]
      },
      {
        "id": "p1203",
        "name": "盖码饭-仔姜牛肉丝",
        "price": 32.0,
        "desc": "仔姜提鲜，牛肉丝滑，鲜辣开胃",
        "tags": [],
        "img": "https://s.koubei.com/xrc8MMFz6xcdYmFZ5fb6sf5kwPrGnEFd.jpg",
        "opts": [
          {
            "g": "现炒码子（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】品牌餐盒",
              "【外带】牛皮纸碗分离",
              "【外带】干湿分离打包盒"
            ]
          }
        ]
      },
      {
        "id": "p1204",
        "name": "盖码饭-酸萝卜炒牛肉",
        "price": 32.0,
        "desc": "酸萝卜爽脆，牛肉丝嫩，微辣够味",
        "tags": [],
        "img": "https://s.koubei.com/83sd4dejX3z6BAY3NWBx8yaiCjkfSrwD.jpg",
        "opts": [
          {
            "g": "现炒码子（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】品牌餐盒",
              "【外带】牛皮纸碗分离",
              "【外带】干湿分离打包盒"
            ]
          }
        ]
      },
      {
        "id": "p1205",
        "name": "盖码饭-回锅牛肉",
        "price": 32.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/fX5Y75XQKmGJSD42rKj24patAZBymGx4.jpg",
        "opts": [
          {
            "g": "现炒码子（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】品牌餐盒",
              "【外带】牛皮纸碗分离",
              "【外带】干湿分离打包盒"
            ]
          }
        ]
      },
      {
        "id": "p1206",
        "name": "盖码饭-卤藕炒回锅牛肉",
        "price": 32.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/R4GyjphebXQ7yXFW4JJwWErCMZ68rnmP.jpg",
        "opts": [
          {
            "g": "现炒码子（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】品牌餐盒",
              "【外带】牛皮纸碗分离",
              "【外带】干湿分离打包盒"
            ]
          }
        ]
      },
      {
        "id": "p1207",
        "name": "盖码饭-四季豆炒牛肚",
        "price": 28.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/8ScGWM45nDP3WtpCAsJnsxr4GsCW5M5s.jpg",
        "opts": [
          {
            "g": "现炒码子（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】品牌餐盒",
              "【外带】牛皮纸碗分离",
              "【外带】干湿分离打包盒"
            ]
          }
        ]
      },
      {
        "id": "p1208",
        "name": "盖码饭-青椒焖牛腩",
        "price": 36.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/3y8S8aahcxBraa4JTSdG8wMcA6bjRM7A.jpg",
        "opts": [
          {
            "g": "现炒码子（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】品牌餐盒",
              "【外带】牛皮纸碗分离",
              "【外带】干湿分离打包盒"
            ]
          }
        ]
      },
      {
        "id": "p1209",
        "name": "盖码饭-樟树港辣椒炒肉",
        "price": 28.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/2wpcKBkHEEx8jSfBRw2dACWeKSTTR5JC.jpg",
        "opts": [
          {
            "g": "现炒码子（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】品牌餐盒",
              "【外带】牛皮纸碗分离",
              "【外带】干湿分离打包盒"
            ]
          }
        ]
      },
      {
        "id": "p1210",
        "name": "盖码饭-香干回锅肉",
        "price": 24.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/jeHw5S5JPE7XBHxYPnXwej4CZwsmQt8D.jpg",
        "opts": [
          {
            "g": "现炒码子（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】品牌餐盒",
              "【外带】牛皮纸碗分离",
              "【外带】干湿分离打包盒"
            ]
          }
        ]
      },
      {
        "id": "p1211",
        "name": "盖码饭-回锅肉炒卤藕",
        "price": 24.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/8tS3ysyETwtyzPStTz8cRJYJHRmTEFwm.jpg",
        "opts": [
          {
            "g": "现炒码子（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】品牌餐盒",
              "【外带】牛皮纸碗分离",
              "【外带】干湿分离打包盒"
            ]
          }
        ]
      },
      {
        "id": "p1212",
        "name": "盖码饭-爆炒顺风",
        "price": 28.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/cDNmQr6SCYacaj4ywZEGtSYxBABWrFGX.jpg",
        "opts": [
          {
            "g": "现炒码子（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】品牌餐盒",
              "【外带】牛皮纸碗分离",
              "【外带】干湿分离打包盒"
            ]
          }
        ]
      },
      {
        "id": "p1213",
        "name": "盖码饭-青椒肉丝",
        "price": 22.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/ajRyNTKjrBXkp2YRe4J4sjhdwpiBpzhH.jpg",
        "opts": [
          {
            "g": "现炒码子（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】品牌餐盒",
              "【外带】牛皮纸碗分离",
              "【外带】干湿分离打包盒"
            ]
          }
        ]
      },
      {
        "id": "p1214",
        "name": "盖码饭-洋葱肉丝",
        "price": 22.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/H33M5kkWbSHXaA5cFB5k3n2yCFdd5FeP.jpg",
        "opts": [
          {
            "g": "现炒码子（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】品牌餐盒",
              "【外带】牛皮纸碗分离",
              "【外带】干湿分离打包盒"
            ]
          }
        ]
      },
      {
        "id": "p1215",
        "name": "盖码饭-榨菜肉丝",
        "price": 22.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/StC2XzKTTfGphjPFtJzmepcJr8PbGPJr.jpg",
        "opts": [
          {
            "g": "现炒码子（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】品牌餐盒",
              "【外带】牛皮纸碗分离",
              "【外带】干湿分离打包盒"
            ]
          }
        ]
      },
      {
        "id": "p1216",
        "name": "盖码饭-鱼香肉丝",
        "price": 28.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/7S3RxNeHCbkfbyyQxXxAnAttKCmFFJPz.jpg",
        "opts": [
          {
            "g": "现炒码子（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】品牌餐盒",
              "【外带】牛皮纸碗分离",
              "【外带】干湿分离打包盒"
            ]
          }
        ]
      },
      {
        "id": "p1217",
        "name": "盖码饭-丝瓜炒肉片",
        "price": 24.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/wxXr8RCQjwezD4Knesxp5mDSptDftnZk.jpg",
        "opts": [
          {
            "g": "现炒码子（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】品牌餐盒",
              "【外带】牛皮纸碗分离",
              "【外带】干湿分离打包盒"
            ]
          }
        ]
      },
      {
        "id": "p1218",
        "name": "盖码饭-爆炒猪肝",
        "price": 22.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/kjcFzJEcQ2ZTfPT35A3kYBmizBGnaaKS.jpg",
        "opts": [
          {
            "g": "现炒码子（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】品牌餐盒",
              "【外带】牛皮纸碗分离",
              "【外带】干湿分离打包盒"
            ]
          }
        ]
      },
      {
        "id": "p1219",
        "name": "盖码饭-爆炒鳝鱼丝",
        "price": 48.0,
        "desc": "鳝片鲜滑，猛火爆炒，微辣鲜香",
        "tags": [],
        "img": "https://s.koubei.com/bj453r2JGy2tXx7T3QJT5MyiCTib6hY3.jpg",
        "opts": [
          {
            "g": "现炒码子（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】品牌餐盒",
              "【外带】牛皮纸碗分离",
              "【外带】干湿分离打包盒"
            ]
          }
        ]
      },
      {
        "id": "p1220",
        "name": "盖码饭-爆炒财鱼片",
        "price": 32.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/kMEik2NsxRjRXBbipEBk8mFktrkZFzs2.jpg",
        "opts": [
          {
            "g": "现炒码子（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】品牌餐盒",
              "【外带】牛皮纸碗分离",
              "【外带】干湿分离打包盒"
            ]
          }
        ]
      },
      {
        "id": "p1221",
        "name": "盖码饭-番茄炒蛋",
        "price": 18.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/RzhdW4QEzwdEfX5BMjfRE8zmCs3npQGi.jpg",
        "opts": [
          {
            "g": "现炒码子（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】品牌餐盒",
              "【外带】牛皮纸碗分离",
              "【外带】干湿分离打包盒"
            ]
          }
        ]
      },
      {
        "id": "p1222",
        "name": "盖码饭-火爆鸡胗",
        "price": 22.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/Zz4ssRWEZHpT3nkANh2pRpy4PBkDr6Ja.jpg",
        "opts": [
          {
            "g": "现炒码子（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】品牌餐盒",
              "【外带】牛皮纸碗分离",
              "【外带】干湿分离打包盒"
            ]
          }
        ]
      },
      {
        "id": "p1223",
        "name": "盖码饭-青椒焖鸡块",
        "price": 28.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/wftXawrZ5ycsMzeCe58CdSi6K7XsRbGh.jpg",
        "opts": [
          {
            "g": "现炒码子（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】品牌餐盒",
              "【外带】牛皮纸碗分离",
              "【外带】干湿分离打包盒"
            ]
          }
        ]
      },
      {
        "id": "p1224",
        "name": "盖码饭-爆炒鲜牛肉",
        "price": 32.0,
        "desc": "鲜牛肉猛火爆炒，滑嫩香辣，锅气十足",
        "tags": [],
        "img": "https://s.koubei.com/2pYzsRt3ABfzx2j7zNmEEpKApY3aHDeS.jpg",
        "opts": [
          {
            "g": "现炒码子（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】品牌餐盒"
            ]
          }
        ]
      },
      {
        "id": "p1225",
        "name": "盖码饭-小炒黄牛肉",
        "price": 32.0,
        "desc": "鲜牛肉猛火爆炒，滑嫩鲜辣，油亮诱人。",
        "tags": [],
        "img": "https://s.koubei.com/SMQx8fTSR73dnfDSiPQDh485wNCkanDh.jpg",
        "opts": [
          {
            "g": "现炒码子（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】品牌餐盒"
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
          "微辣"
        ],
        "img": "https://s.koubei.com/aJt8scATZQiTXHAeeezpfWzM244ijJS8.jpg",
        "opts": [
          {
            "g": "炒粉面饭主食",
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
            "g": "炒粉面饭（默认堂食不用选）",
            "req": true,
            "opts": [
              "【外带】牛皮纸碗",
              "默认堂食碗",
              "【外带】品牌餐盒"
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
          "微辣"
        ],
        "img": "https://s.koubei.com/bf3BRJnZSMxzTkWk8b77MrYM5JzCyhP3.jpg",
        "opts": [
          {
            "g": "炒粉面饭主食",
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
            "g": "炒粉面饭（默认堂食不用选）",
            "req": true,
            "opts": [
              "【外带】牛皮纸碗",
              "默认堂食碗",
              "【外带】品牌餐盒"
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
          "微辣"
        ],
        "img": "https://s.koubei.com/xirnhkYkwCZRJm4AiS5CDHryP7KcXWcC.jpg",
        "opts": [
          {
            "g": "炒粉面饭主食",
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
            "g": "炒粉面饭（默认堂食不用选）",
            "req": true,
            "opts": [
              "【外带】牛皮纸碗",
              "默认堂食碗",
              "【外带】品牌餐盒"
            ]
          }
        ]
      },
      {
        "id": "p1304",
        "name": "（炒粉面饭）卤牛肚",
        "price": 28.0,
        "desc": "卤牛肚嚼劲十足；默认微辣",
        "tags": [
          "微辣"
        ],
        "img": "https://s.koubei.com/X5pDdWJThZFsKxsdcHJemHNr5X3HQEdw.jpg",
        "opts": [
          {
            "g": "炒粉面饭主食",
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
            "g": "炒粉面饭（默认堂食不用选）",
            "req": true,
            "opts": [
              "【外带】牛皮纸碗",
              "默认堂食碗",
              "【外带】品牌餐盒"
            ]
          }
        ]
      },
      {
        "id": "p1305",
        "name": "（炒粉面饭）鸡蛋火腿",
        "price": 18.0,
        "desc": "鸡蛋香嫩，火腿咸香，经典搭配；默认微辣",
        "tags": [
          "微辣"
        ],
        "img": "https://s.koubei.com/5inwRmYiB5663w2TPzhKKiY8bGYBki32.jpg",
        "opts": [
          {
            "g": "炒粉面饭主食",
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
            "g": "炒粉面饭（默认堂食不用选）",
            "req": true,
            "opts": [
              "【外带】牛皮纸碗",
              "默认堂食碗",
              "【外带】品牌餐盒"
            ]
          }
        ]
      },
      {
        "id": "p1306",
        "name": "（炒粉面饭）卤顺风",
        "price": 24.0,
        "desc": "卤猪耳脆嫩嘎嘣脆；默认微辣",
        "tags": [
          "微辣"
        ],
        "img": "https://s.koubei.com/8t5mZYPsmFmHwpfasKsNZGeRx363r6Xw.jpg",
        "opts": [
          {
            "g": "炒粉面饭主食",
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
            "g": "炒粉面饭（默认堂食不用选）",
            "req": true,
            "opts": [
              "【外带】牛皮纸碗",
              "默认堂食碗",
              "【外带】品牌餐盒"
            ]
          }
        ]
      },
      {
        "id": "p1307",
        "name": "（炒粉面饭）回锅肉",
        "price": 24.0,
        "desc": "",
        "tags": [
          "微辣"
        ],
        "img": "https://s.koubei.com/K8adGztWaj4pAij7etrmAMA7T7p6Pyk8.jpg",
        "opts": [
          {
            "g": "炒粉面饭主食",
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
            "g": "炒粉面饭（默认堂食不用选）",
            "req": true,
            "opts": [
              "【外带】牛皮纸碗",
              "默认堂食碗",
              "【外带】品牌餐盒"
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
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/kh22sxhaz6N5PkSJzGQ5JjbXjbHj5tHt.jpg"
      },
      {
        "id": "p1402",
        "name": "原汤粉/面",
        "price": 8.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/QR7SZhNwtQkFY2TKpNMPbrhNA4RPfTBC.jpg",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "圆粉",
              "扁粉",
              "宽粉",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "碱面",
              "豆丝",
              "手工大宽粉"
            ]
          },
          {
            "g": "辣度",
            "req": false,
            "opts": [
              "微辣",
              "清汤不辣"
            ]
          },
          {
            "g": "加码小吃",
            "req": false,
            "opts": [
              "加-卤蛋/枚",
              "加-海带结2个",
              "加-卤鸡爪/只",
              "加-卤鸭掌/只"
            ]
          },
          {
            "g": "粉面（默认堂食不用选）",
            "req": true,
            "opts": [
              "默认堂食碗",
              "【外带】牛皮纸碗",
              "【外带】干湿分离打包盒",
              "【外带】品牌餐盒"
            ]
          }
        ]
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
            "req": true,
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
            "req": true,
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
            "req": true,
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
            "req": true,
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
            "req": true,
            "opts": [
              "常温",
              "冰",
              "温热"
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
            "req": true,
            "opts": [
              "常温",
              "冰",
              "温热"
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
            "req": true,
            "opts": [
              "常温",
              "冰",
              "温热"
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
            "req": true,
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
            "req": true,
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
            "req": true,
            "opts": [
              "常温",
              "冰"
            ]
          }
        ]
      },
      {
        "id": "p1511",
        "name": "大可乐",
        "price": 12.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/wmMi2pzzZN4ZG62Fmdh2ixFnsGfKAdYz.jpg",
        "opts": [
          {
            "g": "温度",
            "req": true,
            "opts": [
              "常温",
              "冰"
            ]
          }
        ]
      },
      {
        "id": "p1512",
        "name": "大雪碧",
        "price": 12.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/jyjYeE7H4hNxT7HJxcrT82QZdfJTaQFC.jpeg",
        "opts": [
          {
            "g": "温度",
            "req": true,
            "opts": [
              "常温",
              "冰"
            ]
          }
        ]
      },
      {
        "id": "p1513",
        "name": "大椰汁",
        "price": 25.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/crYRjmdAwAfNNSxakTDAzeJ52DKS3izE.jpg",
        "opts": [
          {
            "g": "温度",
            "req": true,
            "opts": [
              "常温",
              "冰"
            ]
          }
        ]
      },
      {
        "id": "p1514",
        "name": "瓶装-雪花勇闯天涯",
        "price": 7.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/Y7am2BmHe2e6QnWRsR4hTHwkDnM4MXZr.jpg",
        "opts": [
          {
            "g": "温度",
            "req": true,
            "opts": [
              "常温",
              "冰"
            ]
          }
        ]
      },
      {
        "id": "p1515",
        "name": "小毛铺",
        "price": 20.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/NYX6PH3eKBtm4dXMxAMENB46jxaPxemK.jpg",
        "opts": [
          {
            "g": "温度",
            "req": true,
            "opts": [
              "常温",
              "冰"
            ]
          }
        ]
      },
      {
        "id": "p1516",
        "name": "瓶装-雪花纯生",
        "price": 10.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/8RHBy86ncbRpsEjdsPr4hS6nrhwZaAey.jpg",
        "opts": [
          {
            "g": "温度",
            "req": true,
            "opts": [
              "常温",
              "冰"
            ]
          }
        ]
      },
      {
        "id": "p1517",
        "name": "小劲酒",
        "price": 18.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/2WefBbXW6fF5jixPnRXm7BixzXEXKSdP.jpg",
        "opts": [
          {
            "g": "温度",
            "req": true,
            "opts": [
              "常温",
              "冰"
            ]
          }
        ]
      },
      {
        "id": "p1518",
        "name": "绿豆冰沙",
        "price": 3.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/6yx8GpH33TW37swsmPGNSm52PhMzEwSx.jpg"
      },
      {
        "id": "p1519",
        "name": "维维豆奶-原味",
        "price": 5.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/ZNb4jssTtBWCcWwh4MWfdkxQk8d5N3bW.jpg",
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
        "name": "维维豆奶-绿豆百合味",
        "price": 5.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/dEBKSdxjx462K8hJ2xpb3RbSzmSQBsGb.jpg",
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
        "id": "p1521",
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
        "id": "p1522",
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
  },
  {
    "id": "cat16",
    "name": "打包餐盒",
    "items": [
      {
        "id": "p1601",
        "name": "牛皮纸碗",
        "price": 0.5,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/G62AyzXdCWb6xNXCshMkNkPT55FBhSyZ.jpg"
      },
      {
        "id": "p1602",
        "name": "品牌餐盒",
        "price": 1.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/KHTcAfYwP8Kt6sCyYKWMRihrehdTKAWM.jpg"
      },
      {
        "id": "p1603",
        "name": "品牌餐盒（小）",
        "price": 0.5,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/hHTpaPfHiHHtfES6Fn2dEpFB5hDkFZDA.jpg"
      },
      {
        "id": "p1604",
        "name": "品牌干湿分离打包盒",
        "price": 2.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/eGXCKFRxkwAhQAfMxt847444QkFxirNa.jpg"
      },
      {
        "id": "p1605",
        "name": "炖锅打包盒",
        "price": 2.0,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/EDHNNmkFDsWQyRpXPJjkQj4ebeHyk7iA.jpg"
      },
      {
        "id": "p1606",
        "name": "外卖手提袋",
        "price": 1.5,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/Xi4fhCNaW6krZWYxwJHYJEr7WQepy7Z2.jpg"
      },
      {
        "id": "p1607",
        "name": "码子粉-品牌干湿分离打包盒",
        "price": 2.5,
        "desc": "",
        "tags": [],
        "img": "https://s.koubei.com/ztKDWDQ5fcDBShQRtPcwTie8bdGKbri3.jpg"
      },
      {
        "id": "p1608",
        "name": "鲜米粉/份（纯粉面主食）",
        "price": 3.0,
        "desc": "标准堂食分量，按需购买",
        "tags": [],
        "img": "https://s.koubei.com/hHNNwBjeBYp2J3CGhB3m5wNanCPBYMKA.jpg",
        "opts": [
          {
            "g": "粉面主食",
            "req": true,
            "opts": [
              "圆粉",
              "扁粉",
              "宽粉",
              "手工大宽粉",
              "手工细面",
              "手工扁面",
              "手工宽面",
              "碱面",
              "豆丝"
            ]
          }
        ]
      }
    ]
  }
];
