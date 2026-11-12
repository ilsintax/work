// 통합 통장내역 편집기 기관사전.
// 이 파일을 저장소의 dictionary/institutions.js 와 교체해 올리면 팀 전체에 적용됩니다.
// 아래 { } 안은 순수 JSON 입니다.
(function (root) { root.BANKBOOK_DICTIONARY = {
  "version": 22,
  "updatedAt": "2026-11-12T09:15:11.239Z",
  "source": "reference/stockbalance/index.html BrokerAutoMapper (2026-11-10 추출)",
  "abbreviations": [
    {
      "keys": [
        "KB",
        "국민",
        "KB국민",
        "국민은행",
        "KB국민은행"
      ],
      "bank": "KB국민은행",
      "securities": "KB증권"
    },
    {
      "keys": [
        "신한",
        "신한은행"
      ],
      "bank": "신한은행",
      "securities": "신한투자증권"
    },
    {
      "keys": [
        "NH",
        "농협",
        "NH농협",
        "농협은행"
      ],
      "bank": "NH농협은행",
      "securities": "NH투자증권"
    },
    {
      "keys": [
        "IBK",
        "기업",
        "기업은행"
      ],
      "bank": "IBK기업은행",
      "securities": "IBK투자증권"
    },
    {
      "keys": [
        "IM",
        "iM",
        "하이",
        "대구",
        "iM뱅크",
        "아이엠증권",
        "아이엠"
      ],
      "bank": "iM뱅크",
      "securities": "iM증권"
    },
    {
      "keys": [
        "하나",
        "하나은행"
      ],
      "bank": "하나은행",
      "securities": "하나증권"
    },
    {
      "keys": [
        "우리",
        "우리은행"
      ],
      "bank": "우리은행",
      "securities": "우리투자증권"
    },
    {
      "keys": [
        "SC",
        "제일",
        "SC제일"
      ],
      "bank": "SC제일은행",
      "securities": "한국스탠다드차타드증권"
    },
    {
      "keys": [
        "토스"
      ],
      "bank": "토스뱅크",
      "securities": "토스증권"
    },
    {
      "keys": [
        "카카오"
      ],
      "bank": "카카오뱅크",
      "securities": "카카오페이증권"
    },
    {
      "keys": [
        "미래",
        "미래에셋"
      ],
      "securities": "미래에셋증권"
    },
    {
      "keys": [
        "한투",
        "한국투자"
      ],
      "securities": "한국투자증권"
    },
    {
      "keys": [
        "키움"
      ],
      "securities": "키움증권"
    },
    {
      "keys": [
        "대신"
      ],
      "securities": "대신증권"
    },
    {
      "keys": [
        "교보"
      ],
      "securities": "교보증권"
    },
    {
      "keys": [
        "삼성"
      ],
      "securities": "삼성증권"
    },
    {
      "keys": [
        "DB",
        "디비"
      ],
      "securities": "DB금융투자"
    },
    {
      "keys": [
        "한화"
      ],
      "securities": "한화투자증권"
    },
    {
      "keys": [
        "현대차",
        "현대"
      ],
      "securities": "현대차증권"
    },
    {
      "keys": [
        "메리츠"
      ],
      "securities": "메리츠증권"
    },
    {
      "keys": [
        "유진"
      ],
      "securities": "유진투자증권"
    },
    {
      "keys": [
        "SK"
      ],
      "securities": "SK증권"
    },
    {
      "keys": [
        "LS",
        "이베스트"
      ],
      "securities": "LS증권"
    },
    {
      "keys": [
        "유안타"
      ],
      "securities": "유안타증권"
    },
    {
      "keys": [
        "신영"
      ],
      "securities": "신영증권"
    },
    {
      "keys": [
        "부국"
      ],
      "securities": "부국증권"
    },
    {
      "keys": [
        "다올"
      ],
      "securities": "다올투자증권"
    },
    {
      "keys": [
        "BNK"
      ],
      "securities": "BNK투자증권"
    },
    {
      "keys": [
        "한양"
      ],
      "securities": "한양증권"
    },
    {
      "keys": [
        "케이뱅크"
      ],
      "bank": "케이뱅크"
    },
    {
      "keys": [
        "수협"
      ],
      "bank": "Sh수협은행"
    },
    {
      "keys": [
        "부산"
      ],
      "bank": "부산은행"
    },
    {
      "keys": [
        "경남"
      ],
      "bank": "경남은행"
    },
    {
      "keys": [
        "광주"
      ],
      "bank": "광주은행"
    },
    {
      "keys": [
        "전북"
      ],
      "bank": "전북은행"
    },
    {
      "keys": [
        "제주"
      ],
      "bank": "제주은행"
    },
    {
      "keys": [
        "우체국"
      ],
      "bank": "우체국"
    },
    {
      "keys": [
        "새마을"
      ],
      "bank": "새마을금고"
    },
    {
      "keys": [
        "신협"
      ],
      "bank": "신협"
    },
    {
      "keys": [
        "씨티"
      ],
      "bank": "한국씨티은행"
    }
  ],
  "bank": {
    "columns": {
      "dateCol": [
        "거래일시",
        "거래일자",
        "거래일",
        "일자",
        "날짜",
        "거래 일시"
      ],
      "depositCol": [
        "입금액(원)",
        "입금액",
        "입금",
        "맡기신금액",
        "받으신금액",
        "입금금액"
      ],
      "withdrawalCol": [
        "출금액(원)",
        "출금액",
        "출금",
        "찾으신금액",
        "지급금액",
        "출금금액"
      ],
      "balanceCol": [
        "잔액(원)",
        "잔액",
        "거래후잔액",
        "거래 후 잔액",
        "잔고"
      ],
      "signedAmountCol": [
        "거래금액",
        "거래 금액"
      ],
      "memoCols": [
        "보낸분/받는분",
        "의뢰인/수취인",
        "기재내용",
        "거래내용",
        "받는분/보낸분",
        "내용",
        "적요",
        "내 통장 표시",
        "메모"
      ]
    },
    "profiles": [
      {
        "id": "kbstar",
        "name": "KB국민은행",
        "aliases": [
          "KB국민은행",
          "국민은행"
        ],
        "columns": {
          "dateCol": [
            "거래일시"
          ],
          "depositCol": [
            "입금액(원)"
          ],
          "withdrawalCol": [
            "출금액(원)"
          ],
          "balanceCol": [
            "잔액(원)"
          ],
          "memoCols": [
            "보낸분/받는분"
          ]
        }
      },
      {
        "id": "toss",
        "name": "토스뱅크",
        "aliases": [
          "토스"
        ],
        "columns": {
          "dateCol": [
            "거래 일시"
          ],
          "signedAmountCol": [
            "거래 금액"
          ],
          "balanceCol": [
            "거래 후 잔액"
          ],
          "memoCols": [
            "적요"
          ]
        }
      },
      {
        "id": "custom-12hs165813cw166h1e",
        "name": "신한은행",
        "aliases": [
          "신한은행"
        ],
        "columns": {
          "dateCol": [
            "거래일시"
          ],
          "depositCol": [
            "입금액"
          ],
          "withdrawalCol": [
            "출금액"
          ],
          "balanceCol": [
            "잔액"
          ],
          "memoCols": [
            "내용",
            "적요",
            "비고"
          ]
        },
        "excludedMemoValues": [
          "FB이체",
          "CMS지",
          "서비스",
          "타행FB",
          "타행PC",
          "통지CC",
          "FB자동",
          "FB통신"
        ]
      }
    ],
    "names": [
      "경남은행",
      "광주은행",
      "부산은행",
      "새마을금고",
      "신한은행",
      "신협",
      "우리은행",
      "우체국",
      "전북은행",
      "제주은행",
      "카카오뱅크",
      "케이뱅크",
      "토스뱅크",
      "하나은행",
      "한국씨티은행",
      "IBK기업은행",
      "iM뱅크",
      "KB국민은행",
      "NH농협은행",
      "SC제일은행",
      "Sh수협은행"
    ]
  },
  "securities": {
    "domestic": [
      {
        "id": "kyobo",
        "name": "교보증권",
        "aliases": [
          "교보증권"
        ],
        "profiles": [
          {
            "columns": {
              "dateCol": [
                "거래일자"
              ],
              "tradeTypeCol": [
                "적요명"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "예수금잔고"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "제세금"
              ],
              "stockNameCol": [
                "종목명(거래상대명)"
              ],
              "stockQtyCol": [
                "수량"
              ],
              "unitPriceCol": [
                "단가"
              ],
              "stockBalanceCol": [
                "유가증권잔고"
              ]
            },
            "depositTypes": [
              "정기예탁금이용료입금",
              "펀드환매예약매도"
            ],
            "removalTypes": [
              "공모주입고",
              "타사대체출고"
            ]
          }
        ],
        "shared": {
          "depositTypes": [
            "정기예탁금이용료입금",
            "펀드환매예약매도"
          ],
          "withdrawalTypes": [],
          "removalTypes": [
            "공모주입고",
            "타사대체출고"
          ]
        }
      },
      {
        "id": "daol",
        "name": "다올투자증권",
        "aliases": [
          "다올투자증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "daishin",
        "name": "대신증권",
        "aliases": [
          "대신증권"
        ],
        "profiles": [
          {
            "columns": {
              "dateCol": [
                "거래일자"
              ],
              "tradeTypeCol": [
                "적요명"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "예수금잔고"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "제세금"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "거래수량"
              ],
              "unitPriceCol": [
                "거래단가"
              ],
              "stockBalanceCol": [
                "유가잔고"
              ]
            },
            "depositTypes": [
              "개별상품대체입금",
              "예탁금이용료",
              "이체입금",
              "일괄매도"
            ],
            "withdrawalTypes": [
              "개별상품대체출금",
              "계좌대체출고",
              "공모주청약",
              "기관청약수수료",
              "이체출금",
              "현금매수"
            ],
            "removalTypes": [
              "계좌대체입고",
              "공모주",
              "타사대체입고",
              "타사대체출고",
              "현금매도"
            ],
            "memo": {
              "additionalCol1": [
                "거래상대명"
              ],
              "additionalCol2": [
                "비고"
              ]
            }
          },
          {
            "columns": {
              "dateCol": [
                "거래일자"
              ],
              "tradeTypeCol": [
                "적요명"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "예수금잔액"
              ],
              "feeCol": [
                "세금/수수료"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "거래수량"
              ],
              "unitPriceCol": [
                "채권단가"
              ],
              "stockBalanceCol": [
                "잔고수량"
              ]
            },
            "depositTypes": [
              "개별상품대체입금",
              "예탁금이용료",
              "일괄매도"
            ],
            "withdrawalTypes": [
              "개별상품대체출금",
              "계좌대체출고",
              "현금매수",
              "이체출금"
            ],
            "removalTypes": [
              "공모주",
              "타사대체출고",
              "계좌대체입고",
              "타사대체입고",
              "현금매도"
            ]
          }
        ],
        "shared": {
          "depositTypes": [
            "개별상품대체입금",
            "예탁금이용료",
            "이체입금",
            "일괄매도"
          ],
          "withdrawalTypes": [
            "개별상품대체출금",
            "계좌대체출고",
            "공모주청약",
            "기관청약수수료",
            "이체출금",
            "현금매수"
          ],
          "removalTypes": [
            "계좌대체입고",
            "공모주",
            "타사대체입고",
            "타사대체출고",
            "현금매도"
          ]
        },
        "specialRules": {
          "ignoreTradeTypeContains": [
            "일괄매수"
          ],
          "amountOverrideByTradeType": [
            {
              "match": "일괄매도",
              "sourceHeaders": [
                "이자"
              ]
            }
          ],
          "forceSeparateCosts": true
        }
      },
      {
        "id": "db",
        "name": "DB금융투자",
        "aliases": [
          "DB금융투자",
          "디비금융투자"
        ],
        "profiles": [
          {
            "columns": {
              "dateCol": [
                "일자"
              ],
              "tradeTypeCol": [
                "거래구분"
              ],
              "amountCol": [
                "입출금액"
              ],
              "balanceCol": [
                "예수금"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "세금"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "거래좌수"
              ],
              "unitPriceCol": [
                "기준가"
              ],
              "stockBalanceCol": [
                "잔고좌수"
              ]
            },
            "depositTypes": [
              "예탁금이용료입금",
              "결산분배금입금"
            ],
            "withdrawalTypes": [
              "재투자"
            ],
            "removalTypes": []
          },
          {
            "columns": {
              "dateCol": [
                "거래일자"
              ],
              "tradeTypeCol": [
                "거래구분"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "예수금잔"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "세금"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "거래수량"
              ],
              "unitPriceCol": [
                "거래단가"
              ],
              "stockBalanceCol": [
                "유가금잔"
              ]
            },
            "depositTypes": [
              "예탁금이용료입금",
              "결산분배금입금"
            ],
            "withdrawalTypes": [
              "재투자"
            ],
            "removalTypes": []
          }
        ],
        "shared": {
          "depositTypes": [
            "결산분배금입금",
            "예탁금이용료입금"
          ],
          "withdrawalTypes": [
            "재투자"
          ],
          "removalTypes": []
        }
      },
      {
        "id": "ds",
        "name": "DS투자증권",
        "aliases": [
          "DS투자증권",
          "디에스투자증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "leading",
        "name": "리딩투자증권",
        "aliases": [
          "리딩투자증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "meritz",
        "name": "메리츠증권",
        "aliases": [
          "메리츠증권"
        ],
        "profiles": [
          {
            "columns": {
              "dateCol": [
                "거래일자"
              ],
              "tradeTypeCol": [
                "거래적요"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "예수금잔고"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "제세금"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "수량"
              ],
              "unitPriceCol": [
                "단가"
              ],
              "stockBalanceCol": [
                "유가증권잔고"
              ]
            },
            "removalTypes": [
              "공모주입고",
              "타사대체출고"
            ]
          }
        ],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": [
            "공모주입고",
            "타사대체출고"
          ]
        }
      },
      {
        "id": "mirae",
        "name": "미래에셋증권",
        "aliases": [
          "미래에셋증권"
        ],
        "profiles": [
          {
            "columns": {
              "dateCol": [
                "거래일자"
              ],
              "tradeTypeCol": [
                "거래종류"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "예수금"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "제세금합"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "수량"
              ],
              "unitPriceCol": [
                "단가"
              ],
              "stockBalanceCol": [
                "유가잔고"
              ]
            },
            "depositTypes": [
              "예탁금이용료입금",
              "은행이체입금",
              "주식매도출고"
            ],
            "withdrawalTypes": [
              "계좌대체출금",
              "은행이체대체송금"
            ],
            "removalTypes": [
              "공모주입고",
              "이체입고",
              "이체출고",
              "주식매도입금"
            ],
            "memo": {
              "additionalCol1": [
                "내계좌표시"
              ],
              "additionalCol2": [
                "비고"
              ]
            }
          },
          {
            "columns": {
              "dateCol": [
                "거래일자"
              ],
              "tradeTypeCol": [
                "거래종류"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "예수금잔고"
              ],
              "feeCol": [
                "수수료"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "거래수량"
              ]
            },
            "signature": [
              "거래일자",
              "거래종류",
              "종목명",
              "거래수량",
              "거래금액",
              "외화거래금액",
              "수수료",
              "예수금잔고"
            ],
            "depositTypes": [
              "예탁금이용료입금",
              "은행이체입금"
            ],
            "withdrawalTypes": [
              "계좌대체출금"
            ],
            "removalTypes": [
              "공모주입고",
              "이체출고"
            ],
            "label": "홈페이지 거래내역 간단히 보기(웹페이지)"
          }
        ],
        "shared": {
          "depositTypes": [
            "예탁금이용료입금",
            "은행이체입금",
            "주식매도출고"
          ],
          "withdrawalTypes": [
            "계좌대체출금",
            "은행이체대체송금"
          ],
          "removalTypes": [
            "공모주입고",
            "이체입고",
            "이체출고",
            "주식매도입금"
          ]
        }
      },
      {
        "id": "bookook",
        "name": "부국증권",
        "aliases": [
          "부국증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "bnk",
        "name": "BNK투자증권",
        "aliases": [
          "BNK투자증권",
          "비엔케이투자증권"
        ],
        "profiles": [
          {
            "columns": {
              "dateCol": [
                "일자"
              ],
              "tradeTypeCol": [
                "거래구분"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "예수금잔"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "거래세"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "거래수량"
              ],
              "unitPriceCol": [
                "단가"
              ],
              "stockBalanceCol": [
                "유가금잔"
              ]
            },
            "removalTypes": []
          }
        ],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "samsung",
        "name": "삼성증권",
        "aliases": [
          "삼성증권"
        ],
        "profiles": [
          {
            "columns": {
              "dateCol": [
                "거래일자"
              ],
              "tradeTypeCol": [
                "거래명"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "현금잔액"
              ],
              "feeCol": [
                "수수료/Fee"
              ],
              "taxCol": [
                "제세금/대출이자"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "거래수량"
              ],
              "unitPriceCol": [
                "거래단가"
              ],
              "stockBalanceCol": [
                "잔고수량/펀드평가금액"
              ]
            },
            "depositTypes": [
              "이용료입금"
            ],
            "removalTypes": [
              "청약입고",
              "타사출고"
            ]
          },
          {
            "columns": {
              "dateCol": [
                "거래일자"
              ],
              "tradeTypeCol": [
                "거래명"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "현금잔액"
              ],
              "feeCol": [
                "수수료/Fee"
              ],
              "taxCol": [
                "제세금/대출이자"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "거래수량"
              ],
              "unitPriceCol": [
                "거래단가"
              ],
              "stockBalanceCol": [
                "잔고수량"
              ]
            },
            "depositTypes": [
              "이용료입금"
            ],
            "removalTypes": [
              "청약입고",
              "타사출고"
            ]
          }
        ],
        "shared": {
          "depositTypes": [
            "이용료입금"
          ],
          "withdrawalTypes": [],
          "removalTypes": [
            "청약입고",
            "타사출고"
          ]
        }
      },
      {
        "id": "sangsangin",
        "name": "상상인증권",
        "aliases": [
          "상상인증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "shinyoung",
        "name": "신영증권",
        "aliases": [
          "신영증권"
        ],
        "profiles": [
          {
            "columns": {
              "dateCol": [
                "거래일자"
              ],
              "tradeTypeCol": [
                "거래적요"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "RP+예수금"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "제세금"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "수량"
              ],
              "unitPriceCol": [
                "단가"
              ],
              "stockBalanceCol": [
                "잔량"
              ]
            },
            "removalTypes": [
              "타사이체출고"
            ]
          }
        ],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": [
            "타사이체출고"
          ]
        }
      },
      {
        "id": "shinhan",
        "name": "신한투자증권",
        "aliases": [
          "신한투자증권"
        ],
        "profiles": [
          {
            "columns": {
              "dateCol": [
                "일자"
              ],
              "tradeTypeCol": [
                "구분"
              ],
              "amountCol": [
                "거래대금"
              ],
              "balanceCol": [
                "최종금액"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "제세금"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "수량"
              ],
              "unitPriceCol": [
                "가격"
              ]
            },
            "depositTypes": [
              "배당금",
              "예탁금이용료",
              "장내_매도",
              "코스닥_매도",
              "환전입금"
            ],
            "withdrawalTypes": [
              "은행이체출금",
              "장내_매수",
              "코스닥_매수"
            ],
            "removalTypes": [
              "공모주입고",
              "타사대체출고"
            ],
            "memo": {
              "additionalCol1": [
                "의뢰자명"
              ]
            }
          },
          {
            "columns": {
              "dateCol": [
                "일자"
              ],
              "tradeTypeCol": [
                "구분"
              ],
              "amountCol": [
                "거래대금"
              ],
              "balanceCol": [
                "최종금액"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "제세금"
              ],
              "stockNameCol": [
                "종목명(상대처)"
              ],
              "stockQtyCol": [
                "수량"
              ],
              "unitPriceCol": [
                "가격"
              ]
            },
            "depositTypes": [
              "예탁금이용료",
              "장내_매도",
              "코스닥_매도",
              "환전입금",
              "배당금"
            ],
            "withdrawalTypes": [
              "장내_매수",
              "코스닥_매수",
              "은행이체출금"
            ],
            "removalTypes": [
              "공모주입고",
              "타사대체출고"
            ]
          }
        ],
        "shared": {
          "depositTypes": [
            "배당금",
            "예탁금이용료",
            "장내_매도",
            "코스닥_매도",
            "환전입금"
          ],
          "withdrawalTypes": [
            "은행이체출금",
            "장내_매수",
            "코스닥_매수"
          ],
          "removalTypes": [
            "공모주입고",
            "타사대체출고"
          ]
        }
      },
      {
        "id": "si",
        "name": "SI증권",
        "aliases": [
          "SI증권",
          "에스아이증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "sk",
        "name": "SK증권",
        "aliases": [
          "SK증권",
          "에스케이증권"
        ],
        "profiles": [
          {
            "columns": {
              "dateCol": [
                "거래일"
              ],
              "tradeTypeCol": [
                "거래종류"
              ],
              "amountCol": [
                "세전금액"
              ],
              "balanceCol": [
                "예수금잔고"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "거래세+농특세"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "수량"
              ],
              "unitPriceCol": [
                "단가"
              ],
              "stockBalanceCol": [
                "유가증권잔고"
              ]
            },
            "depositTypes": [
              "예탁금이용료(원화)",
              "사채원금"
            ],
            "withdrawalTypes": [
              "이체-출금"
            ],
            "removalTypes": []
          },
          {
            "columns": {
              "dateCol": [
                "거래일자"
              ],
              "tradeTypeCol": [
                "거래종류"
              ],
              "amountCol": [
                "정산금액"
              ],
              "balanceCol": [
                "예수금잔고"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "소득/주민/법인세"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "수량"
              ],
              "unitPriceCol": [
                "단가"
              ],
              "stockBalanceCol": [
                "유가증권잔고"
              ]
            },
            "signature": [
              "거래일자",
              "거래종류",
              "종목코드",
              "수량",
              "정산금액",
              "예수금잔고",
              "미수미납변제:",
              "수수료",
              "소득/주민/법인세",
              "연체료",
              "상품위험등급",
              "금융기관명",
              "비고",
              "거래번호",
              "신용종류",
              "종목명",
              "단가",
              "제비용합",
              "유가증권잔고",
              "미수발생금",
              "거래세농특세",
              "신용이자",
              "기초자산/신용등급",
              "상대계좌번호"
            ],
            "depositTypes": [
              "이체-입금"
            ],
            "withdrawalTypes": [
              "공모주 기관청약 납입",
              "타사대체-출고"
            ],
            "removalTypes": [
              "공모주"
            ],
            "memo": {
              "additionalCol1": [
                "비고"
              ]
            },
            "inferredFeeMode": "included"
          }
        ],
        "shared": {
          "depositTypes": [
            "사채원금",
            "예탁금이용료(원화)",
            "이체-입금"
          ],
          "withdrawalTypes": [
            "공모주 기관청약 납입",
            "이체-출금",
            "타사대체-출고"
          ],
          "removalTypes": [
            "공모주"
          ]
        }
      },
      {
        "id": "nh",
        "name": "NH투자증권",
        "aliases": [
          "NH투자증권",
          "엔에이치투자증권"
        ],
        "profiles": [
          {
            "columns": {
              "dateCol": [
                "실거래일자"
              ],
              "tradeTypeCol": [
                "거래종류"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "잔고금액"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "세금"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "수량"
              ],
              "unitPriceCol": [
                "가격"
              ],
              "stockBalanceCol": [
                "잔고수량"
              ]
            },
            "depositTypes": [
              "KOSDAQ매도",
              "예탁금이용료",
              "은행이체입금"
            ],
            "withdrawalTypes": [
              "IPO청약수수료출금",
              "공모청약출금",
              "업무수수료출금"
            ],
            "removalTypes": [
              "공모주입고",
              "타사대체출고"
            ],
            "memo": {
              "additionalCol1": [
                "받는통장 표시내용"
              ],
              "additionalCol2": [
                "비고"
              ]
            }
          },
          {
            "columns": {
              "dateCol": [
                "실거래일자"
              ],
              "tradeTypeCol": [
                "상세내용"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "잔고금액"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "세금"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "수량"
              ],
              "unitPriceCol": [
                "단가"
              ],
              "stockBalanceCol": [
                "잔고"
              ]
            },
            "depositTypes": [
              "은행이체입금",
              "KOSDAQ매도",
              "예탁금이용료"
            ],
            "withdrawalTypes": [
              "IPO청약수수료출금",
              "공모청약출금",
              "업무수수료출금"
            ],
            "removalTypes": [
              "공모주입고",
              "타사대체출고"
            ]
          },
          {
            "columns": {
              "dateCol": [
                "실거래일자"
              ],
              "tradeTypeCol": [
                "거래종류"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "잔고금액"
              ]
            },
            "depositTypes": [
              "은행이체입금",
              "KOSDAQ매도",
              "예탁금이용료"
            ],
            "withdrawalTypes": [
              "IPO청약수수료출금",
              "공모청약출금",
              "업무수수료출금"
            ],
            "removalTypes": [
              "공모주입고",
              "타사대체출고"
            ],
            "minMatchedKeys": 4
          }
        ],
        "shared": {
          "depositTypes": [
            "KOSDAQ매도",
            "예탁금이용료",
            "은행이체입금"
          ],
          "withdrawalTypes": [
            "IPO청약수수료출금",
            "공모청약출금",
            "업무수수료출금"
          ],
          "removalTypes": [
            "공모주입고",
            "타사대체출고"
          ]
        }
      },
      {
        "id": "ls",
        "name": "LS증권",
        "aliases": [
          "LS증권",
          "엘에스증권",
          "이베스트투자증권"
        ],
        "profiles": [
          {
            "columns": {
              "dateCol": [
                "결제일자"
              ],
              "tradeTypeCol": [
                "거래종류"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "예수금잔고"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "제세금합"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "거래수량"
              ],
              "unitPriceCol": [
                "거래단가"
              ],
              "stockBalanceCol": [
                "유가잔고"
              ]
            },
            "depositTypes": [
              "KOSDAQ보통매도_KRX",
              "예탁금이용료입금"
            ],
            "withdrawalTypes": [
              "은행이체출금"
            ],
            "removalTypes": [
              "타사대체입고"
            ],
            "memo": {
              "additionalCol1": [
                "상대계좌명"
              ],
              "additionalCol2": [
                "비고"
              ]
            }
          }
        ],
        "shared": {
          "depositTypes": [
            "KOSDAQ보통매도_KRX",
            "예탁금이용료입금"
          ],
          "withdrawalTypes": [
            "은행이체출금"
          ],
          "removalTypes": [
            "타사대체입고"
          ]
        }
      },
      {
        "id": "woori",
        "name": "우리투자증권",
        "aliases": [
          "우리투자증권",
          "한국포스증권",
          "우리종합금융"
        ],
        "profiles": [
          {
            "columns": {
              "dateCol": [
                "거래일자"
              ],
              "tradeTypeCol": [
                "거래구분"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "예수금"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "제세금"
              ],
              "stockNameCol": [
                "펀드명"
              ],
              "stockQtyCol": [
                "거래좌수"
              ],
              "unitPriceCol": [
                "기준가"
              ],
              "stockBalanceCol": [
                "잔고좌수"
              ]
            },
            "removalTypes": []
          }
        ],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "yuanta",
        "name": "유안타증권",
        "aliases": [
          "유안타증권"
        ],
        "profiles": [
          {
            "columns": {
              "dateCol": [
                "거래일"
              ],
              "tradeTypeCol": [
                "적요"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "예수금 잔액"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "제세금(거래세)"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "거래수량"
              ],
              "unitPriceCol": [
                "거래단가"
              ],
              "stockBalanceCol": [
                "잔 고"
              ]
            },
            "removalTypes": []
          }
        ],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "eugene",
        "name": "유진투자증권",
        "aliases": [
          "유진투자증권"
        ],
        "profiles": [
          {
            "columns": {
              "dateCol": [
                "일자"
              ],
              "tradeTypeCol": [
                "거래구분"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "잔액"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "제세금"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "수량"
              ],
              "unitPriceCol": [
                "매매가격"
              ],
              "stockBalanceCol": [
                "잔고"
              ]
            },
            "depositTypes": [
              "공모불입(현금)",
              "입금",
              "정기예탁금이용료"
            ],
            "withdrawalTypes": [
              "공모주 추가납입",
              "청약수수료 징수"
            ],
            "removalTypes": [
              "출고",
              "공모주",
              "공모불입(현금)"
            ]
          },
          {
            "columns": {
              "dateCol": [
                "거래일자"
              ],
              "tradeTypeCol": [
                "거래종류"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "예수금잔고"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "제세금"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "수량"
              ],
              "unitPriceCol": [
                "단가"
              ],
              "stockBalanceCol": [
                "유가증권잔고"
              ]
            },
            "depositTypes": [
              "공모불입(현금)",
              "입금",
              "정기예탁금이용료"
            ],
            "withdrawalTypes": [
              "공모주 추가납입",
              "청약수수료 징수"
            ],
            "removalTypes": [
              "출고",
              "공모주",
              "공모불입(현금)"
            ]
          }
        ],
        "shared": {
          "depositTypes": [
            "입금",
            "정기예탁금이용료"
          ],
          "withdrawalTypes": [
            "공모주 추가납입",
            "청약수수료 징수"
          ],
          "removalTypes": [
            "공모주",
            "출고"
          ]
        }
      },
      {
        "id": "yuhwa",
        "name": "유화증권",
        "aliases": [
          "유화증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "im",
        "name": "iM증권",
        "aliases": [
          "iM증권",
          "아이엠증권",
          "하이투자증권"
        ],
        "profiles": [
          {
            "columns": {
              "dateCol": [
                "거래일자"
              ],
              "tradeTypeCol": [
                "적요명"
              ],
              "amountCol": [
                "거래금액"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "제세금"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "거래수량"
              ],
              "unitPriceCol": [
                "단가"
              ],
              "directionCol": [
                "거래구분"
              ]
            },
            "signature": [
              "거래일자",
              "번호",
              "처리지점",
              "거래구분",
              "적요명",
              "종목명",
              "단가",
              "거래수량",
              "거래금액",
              "수수료",
              "제세금",
              "비고"
            ],
            "depositTypes": [
              "이체,입금"
            ],
            "withdrawalTypes": [
              "공모주기관납입출금"
            ],
            "removalTypes": [],
            "noBalance": true,
            "label": "거래내역(잔액 없음)",
            "inferredFeeMode": "included",
            "memo": {
              "additionalCol1": [
                "비고"
              ]
            }
          }
        ],
        "shared": {
          "depositTypes": [
            "이체,입금"
          ],
          "withdrawalTypes": [
            "공모주기관납입출금"
          ],
          "removalTypes": []
        }
      },
      {
        "id": "ibk",
        "name": "IBK투자증권",
        "aliases": [
          "IBK투자증권",
          "아이비케이투자증권"
        ],
        "profiles": [
          {
            "columns": {
              "dateCol": [
                "결제일자"
              ],
              "tradeTypeCol": [
                "거래종류"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "예수금잔고"
              ],
              "feeCol": [
                "거래수수료"
              ],
              "taxCol": [
                "제세금합"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "거래수량"
              ],
              "unitPriceCol": [
                "거래단가"
              ],
              "stockBalanceCol": [
                "유가잔고"
              ]
            },
            "depositTypes": [
              "예탁금이용료입금"
            ],
            "removalTypes": [
              "공모주입고",
              "타사대체출고"
            ]
          }
        ],
        "shared": {
          "depositTypes": [
            "예탁금이용료입금"
          ],
          "withdrawalTypes": [],
          "removalTypes": [
            "공모주입고",
            "타사대체출고"
          ]
        }
      },
      {
        "id": "kakaopay",
        "name": "카카오페이증권",
        "aliases": [
          "카카오페이증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "kb",
        "name": "KB증권",
        "aliases": [
          "KB증권",
          "케이비증권"
        ],
        "profiles": [
          {
            "columns": {
              "dateCol": [
                "거래일자"
              ],
              "tradeTypeCol": [
                "거래종류"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "예수금"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "거래세 등"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "수량"
              ],
              "unitPriceCol": [
                "단가"
              ],
              "stockBalanceCol": [
                "유가잔고"
              ]
            },
            "depositTypes": [
              "은행이체 입금"
            ],
            "withdrawalTypes": [
              "기관납입 출금",
              "타사대체수수료 출금"
            ],
            "removalTypes": [
              "공모주 입고",
              "타사대체 출고"
            ],
            "inferredFeeMode": "included",
            "memo": {
              "additionalCol1": [
                "비고"
              ]
            }
          },
          {
            "columns": {
              "dateCol": [
                "거래일"
              ],
              "tradeTypeCol": [
                "내용"
              ],
              "balanceCol": [
                "예수금잔액(원)"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "수량"
              ],
              "unitPriceCol": [
                "단가"
              ],
              "kbDepositAmountCol": [
                "입금/입고/매도",
                "입금 / 입고 / 매도"
              ],
              "kbWithdrawalAmountCol": [
                "출금/출고/매수",
                "출금 / 출고 / 매수"
              ]
            },
            "depositTypes": [
              "은행이체 입금"
            ],
            "withdrawalTypes": [
              "기관납입 출금",
              "타사대체수수료출금"
            ],
            "removalTypes": [
              "공모주 입고",
              "타사대체 출고"
            ],
            "specialRules": {
              "deriveAmountFromTradeTypeText": true,
              "classificationKeywords": {
                "deposit": [
                  "입금",
                  "입고",
                  "매도"
                ],
                "withdrawal": [
                  "출금",
                  "출고",
                  "매수"
                ]
              },
              "unmatchedClassification": "remove",
              "notice": "KB증권 간편 포맷으로 인식되었습니다. 거래금액 열이 없어도 진행되며, 입금/입고/매도 또는 출금/출고/매수 컬럼의 금액만 거래금액으로 사용합니다. 잔액 차이로 금액을 자동 추정하지 않습니다."
            }
          }
        ],
        "shared": {
          "depositTypes": [
            "은행이체 입금"
          ],
          "withdrawalTypes": [
            "기관납입 출금",
            "타사대체수수료 출금"
          ],
          "removalTypes": [
            "공모주 입고",
            "타사대체 출고"
          ]
        }
      },
      {
        "id": "kr",
        "name": "KR투자증권",
        "aliases": [
          "KR투자증권",
          "케이알투자증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "cape",
        "name": "케이프투자증권",
        "aliases": [
          "케이프투자증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "koreaasset",
        "name": "코리아에셋투자증권",
        "aliases": [
          "코리아에셋투자증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "kiwoom",
        "name": "키움증권",
        "aliases": [
          "키움증권"
        ],
        "profiles": [
          {
            "columns": {
              "dateCol": [
                "거래일자"
              ],
              "tradeTypeCol": [
                "거래종류"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "예수금"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "거래세/농특세"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "거래수량"
              ],
              "unitPriceCol": [
                "거래단가"
              ],
              "stockBalanceCol": [
                "유가잔고"
              ]
            },
            "depositTypes": [
              "예탁금이용료(이자)입금",
              "자금부입금법인청약",
              "예탁금이용료(이자)",
              "배당금입금",
              "수익분배금입금",
              "이체입금(연계은행)",
              "장내매도",
              "세금외화매도"
            ],
            "withdrawalTypes": [
              "공모불입출금",
              "수요예측청약수수료출금",
              "이체출금",
              "장내매수",
              "예탁금이용료이자세금출금",
              "KOSDAQ매수"
            ],
            "removalTypes": [
              "공모주입고",
              "타사대체출고"
            ]
          },
          {
            "columns": {
              "dateCol": [
                "거래일자"
              ],
              "tradeTypeCol": [
                "거래구분"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "예수금"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "거래세/농특세"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "거래수량"
              ],
              "unitPriceCol": [
                "거래단가"
              ],
              "stockBalanceCol": [
                "유가잔고"
              ]
            },
            "depositTypes": [
              "예탁금이용료(이자)입금",
              "자금부입금법인청약",
              "예탁금이용료(이자)",
              "배당금입금",
              "수익분배금입금",
              "이체입금(연계은행)",
              "장내매도",
              "세금외화매도"
            ],
            "withdrawalTypes": [
              "공모불입출금",
              "수요예측청약수수료출금",
              "이체출금",
              "장내매수",
              "예탁금이용료이자세금출금",
              "KOSDAQ매수"
            ],
            "removalTypes": [
              "공모주입고",
              "타사대체출고"
            ]
          },
          {
            "columns": {
              "dateCol": [
                "일자"
              ],
              "tradeTypeCol": [
                "거래구분"
              ],
              "amountCol": [
                "입출금액"
              ],
              "balanceCol": [
                "예수금잔고"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "제세금"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "거래좌수"
              ],
              "unitPriceCol": [
                "기준가"
              ],
              "stockBalanceCol": [
                "잔고좌수"
              ]
            },
            "depositTypes": [
              "예탁금이용료(이자)입금",
              "자금부입금법인청약",
              "예탁금이용료(이자)",
              "배당금입금",
              "수익분배금입금",
              "이체입금(연계은행)",
              "장내매도",
              "세금외화매도"
            ],
            "withdrawalTypes": [
              "공모불입출금",
              "수요예측청약수수료출금",
              "이체출금",
              "장내매수",
              "예탁금이용료이자세금출금",
              "KOSDAQ매수"
            ],
            "removalTypes": [
              "공모주입고",
              "타사대체출고"
            ]
          },
          {
            "columns": {
              "dateCol": [
                "거래일자"
              ],
              "tradeTypeCol": [
                "거래구분"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "예수금"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "거래세/농특세"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "거래수량"
              ],
              "unitPriceCol": [
                "거래단가"
              ],
              "stockBalanceCol": [
                "유가잔고"
              ]
            },
            "depositTypes": [
              "예탁금이용료(이자)입금",
              "자금부입금법인청약",
              "예탁금이용료(이자)",
              "배당금입금",
              "수익분배금입금",
              "이체입금(연계은행)",
              "장내매도",
              "세금외화매도"
            ],
            "withdrawalTypes": [
              "공모불입출금",
              "수요예측청약수수료출금",
              "이체출금",
              "장내매수",
              "예탁금이용료이자세금출금",
              "KOSDAQ매수"
            ],
            "removalTypes": [
              "공모주입고",
              "타사대체출고"
            ]
          }
        ],
        "shared": {
          "depositTypes": [
            "배당금입금",
            "세금외화매도",
            "수익분배금입금",
            "예탁금이용료(이자)",
            "예탁금이용료(이자)입금",
            "이체입금(연계은행)",
            "자금부입금법인청약",
            "장내매도"
          ],
          "withdrawalTypes": [
            "KOSDAQ매수",
            "공모불입출금",
            "수요예측청약수수료출금",
            "예탁금이용료이자세금출금",
            "이체출금",
            "장내매수"
          ],
          "removalTypes": [
            "공모주입고",
            "타사대체출고"
          ]
        }
      },
      {
        "id": "toss",
        "name": "토스증권",
        "aliases": [
          "토스증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "hana",
        "name": "하나증권",
        "aliases": [
          "하나증권"
        ],
        "profiles": [
          {
            "columns": {
              "dateCol": [
                "거래일자"
              ],
              "tradeTypeCol": [
                "거래구분"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "예수금잔액"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "거래/소득세"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "거래수량"
              ],
              "unitPriceCol": [
                "거래단가"
              ],
              "stockBalanceCol": [
                "유가증권금잔"
              ]
            },
            "depositTypes": [
              "예탁금이용료"
            ],
            "removalTypes": [
              "타사대체출고",
              "공모주입고"
            ]
          }
        ],
        "shared": {
          "depositTypes": [
            "예탁금이용료"
          ],
          "withdrawalTypes": [],
          "removalTypes": [
            "공모주입고",
            "타사대체출고"
          ]
        }
      },
      {
        "id": "sc",
        "name": "한국스탠다드차타드증권",
        "aliases": [
          "한국스탠다드차타드증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "koreainvest",
        "name": "한국투자증권",
        "aliases": [
          "한국투자증권"
        ],
        "profiles": [
          {
            "columns": {
              "dateCol": [
                "거래일"
              ],
              "tradeTypeCol": [
                "거래종류"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "예수금"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "거래세"
              ],
              "extraCostCol1": [
                "세금"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "거래수량"
              ],
              "unitPriceCol": [
                "거래단가"
              ],
              "stockBalanceCol": [
                "유가잔고"
              ]
            },
            "depositTypes": [
              "HTS코스닥주식매도",
              "예탁금이용료",
              "타사이체입금",
              "펀드현금배당"
            ],
            "withdrawalTypes": [
              "HTS타사이체출금",
              "WTS추납대체청약",
              "기관일괄대체청약납입",
              "출고수수료"
            ],
            "removalTypes": [
              "공모주입고",
              "타사이체출고"
            ],
            "memo": {
              "additionalCol1": [
                "비고"
              ]
            }
          },
          {
            "columns": {
              "dateCol": [
                "거래일"
              ],
              "tradeTypeCol": [
                "거래종류"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "잔액"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "거래세"
              ],
              "extraCostCol1": [
                "세금"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "거래수량"
              ],
              "unitPriceCol": [
                "거래단가"
              ],
              "stockBalanceCol": [
                "유가잔고"
              ]
            },
            "depositTypes": [
              "HTS코스닥주식매도",
              "타사이체입금",
              "예탁금이용료",
              "펀드현금배당"
            ],
            "withdrawalTypes": [
              "HTS타사이체출금",
              "WTS추납대체청약",
              "기관일괄대체청약납입",
              "출고수수료"
            ],
            "removalTypes": [
              "공모주입고",
              "타사이체출고"
            ]
          }
        ],
        "shared": {
          "depositTypes": [
            "HTS코스닥주식매도",
            "예탁금이용료",
            "타사이체입금",
            "펀드현금배당"
          ],
          "withdrawalTypes": [
            "HTS타사이체출금",
            "WTS추납대체청약",
            "기관일괄대체청약납입",
            "출고수수료"
          ],
          "removalTypes": [
            "공모주입고",
            "타사이체출고"
          ]
        }
      },
      {
        "id": "hanyang",
        "name": "한양증권",
        "aliases": [
          "한양증권"
        ],
        "profiles": [
          {
            "columns": {
              "dateCol": [
                "일자"
              ],
              "tradeTypeCol": [
                "거래구분"
              ],
              "amountCol": [
                "입출구분"
              ],
              "balanceCol": [
                "예수금잔"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "제세금"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "거래좌수"
              ],
              "unitPriceCol": [
                "기준가"
              ],
              "stockBalanceCol": [
                "잔고좌수"
              ]
            },
            "depositTypes": [
              "결산분배금입금"
            ],
            "withdrawalTypes": [
              "은행이체출금"
            ],
            "removalTypes": []
          }
        ],
        "shared": {
          "depositTypes": [
            "결산분배금입금"
          ],
          "withdrawalTypes": [
            "은행이체출금"
          ],
          "removalTypes": []
        }
      },
      {
        "id": "hanhwa",
        "name": "한화투자증권",
        "aliases": [
          "한화투자증권"
        ],
        "profiles": [
          {
            "columns": {
              "dateCol": [
                "거래일"
              ],
              "tradeTypeCol": [
                "거래구분"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "예수금"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "제세금합"
              ],
              "stockNameCol": [
                "종목(상품)명"
              ],
              "stockQtyCol": [
                "수량"
              ],
              "unitPriceCol": [
                "단가"
              ]
            },
            "depositTypes": [
              "[입금] 예탁금이용료입금"
            ],
            "removalTypes": []
          },
          {
            "columns": {
              "dateCol": [
                "거래일자"
              ],
              "tradeTypeCol": [
                "거래구분"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "원화예수금"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "제세금합"
              ],
              "stockNameCol": [
                "종목(상품)명"
              ],
              "stockQtyCol": [
                "수량"
              ],
              "unitPriceCol": [
                "단가"
              ]
            },
            "depositTypes": [
              "[입금] 예탁금이용료입금"
            ],
            "removalTypes": []
          },
          {
            "columns": {
              "dateCol": [
                "거래일자"
              ],
              "tradeTypeCol": [
                "거래구분"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "원화예수금"
              ],
              "feeCol": [
                "수수료(VAT포함)"
              ],
              "taxCol": [
                "제세금합"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "수량"
              ],
              "unitPriceCol": [
                "단가"
              ],
              "stockBalanceCol": [
                "잔고수량"
              ]
            },
            "signature": [
              "거래일자",
              "거래NO",
              "거래구분",
              "수량",
              "거래금액",
              "수수료(VAT포함)",
              "상환/변제금액",
              "원화예수금",
              "미수발생금",
              "비고",
              "계약번호",
              "취소여부",
              "원거래NO",
              "종목명",
              "단가",
              "정산금액",
              "제세금합",
              "융자/대출이자",
              "외화예수금",
              "잔고수량"
            ],
            "depositTypes": [
              "[입금] 은행이체입금(…)"
            ],
            "withdrawalTypes": [
              "[출금] 청약증거금출금"
            ],
            "removalTypes": [
              "[입고] 공모주입고",
              "[출고] 타사대체출고(…)"
            ],
            "memo": {
              "additionalCol1": [
                "비고"
              ]
            }
          }
        ],
        "shared": {
          "depositTypes": [
            "[입금] 예탁금이용료입금",
            "[입금] 은행이체입금(…)"
          ],
          "withdrawalTypes": [
            "[출금] 청약증거금출금"
          ],
          "removalTypes": [
            "[입고] 공모주입고",
            "[출고] 타사대체출고(…)"
          ]
        }
      },
      {
        "id": "hyundai",
        "name": "현대차증권",
        "aliases": [
          "현대차증권"
        ],
        "profiles": [
          {
            "columns": {
              "dateCol": [
                "일자"
              ],
              "tradeTypeCol": [
                "거래구분"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "예수금잔(원)"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "세금합"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "거래수량"
              ],
              "unitPriceCol": [
                "단가"
              ],
              "stockBalanceCol": [
                "유가금잔"
              ]
            },
            "withdrawalTypes": [
              "매수"
            ],
            "removalTypes": []
          },
          {
            "columns": {
              "dateCol": [
                "거래일자"
              ],
              "tradeTypeCol": [
                "거래구분"
              ],
              "amountCol": [
                "거래금액"
              ],
              "balanceCol": [
                "예수금잔"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "세금합"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "거래좌수"
              ],
              "unitPriceCol": [
                "기준가"
              ],
              "stockBalanceCol": [
                "잔고좌수"
              ]
            },
            "withdrawalTypes": [
              "매수"
            ],
            "removalTypes": []
          }
        ],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [
            "매수"
          ],
          "removalTypes": []
        }
      },
      {
        "id": "heungkuk",
        "name": "흥국증권",
        "aliases": [
          "흥국증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      }
    ],
    "foreign": [
      {
        "id": "kyobo",
        "name": "교보증권",
        "aliases": [
          "교보증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "daol",
        "name": "다올투자증권",
        "aliases": [
          "다올투자증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "daishin",
        "name": "대신증권",
        "aliases": [
          "대신증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "db",
        "name": "DB금융투자",
        "aliases": [
          "DB금융투자",
          "디비금융투자"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "ds",
        "name": "DS투자증권",
        "aliases": [
          "DS투자증권",
          "디에스투자증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "leading",
        "name": "리딩투자증권",
        "aliases": [
          "리딩투자증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "meritz",
        "name": "메리츠증권",
        "aliases": [
          "메리츠증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "mirae",
        "name": "미래에셋증권",
        "aliases": [
          "미래에셋증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "bookook",
        "name": "부국증권",
        "aliases": [
          "부국증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "bnk",
        "name": "BNK투자증권",
        "aliases": [
          "BNK투자증권",
          "비엔케이투자증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "samsung",
        "name": "삼성증권",
        "aliases": [
          "삼성증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "sangsangin",
        "name": "상상인증권",
        "aliases": [
          "상상인증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "shinyoung",
        "name": "신영증권",
        "aliases": [
          "신영증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "shinhan",
        "name": "신한투자증권",
        "aliases": [
          "신한투자증권"
        ],
        "profiles": [
          {
            "columns": {
              "dateCol": [
                "일자"
              ],
              "tradeTypeCol": [
                "구분"
              ],
              "amountCol": [
                "거래대금"
              ],
              "balanceCol": [
                "최종금액"
              ],
              "feeCol": [
                "수수료"
              ],
              "taxCol": [
                "제세금"
              ],
              "stockNameCol": [
                "종목명(상대처)"
              ],
              "stockQtyCol": [
                "수량"
              ],
              "unitPriceCol": [
                "가격"
              ]
            },
            "withdrawalTypes": [
              "환전출금"
            ]
          }
        ],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [
            "환전출금"
          ],
          "removalTypes": []
        }
      },
      {
        "id": "si",
        "name": "SI증권",
        "aliases": [
          "SI증권",
          "에스아이증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "sk",
        "name": "SK증권",
        "aliases": [
          "SK증권",
          "에스케이증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "nh",
        "name": "NH투자증권",
        "aliases": [
          "NH투자증권",
          "엔에이치투자증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "ls",
        "name": "LS증권",
        "aliases": [
          "LS증권",
          "엘에스증권",
          "이베스트투자증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "woori",
        "name": "우리투자증권",
        "aliases": [
          "우리투자증권",
          "한국포스증권",
          "우리종합금융"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "yuanta",
        "name": "유안타증권",
        "aliases": [
          "유안타증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "eugene",
        "name": "유진투자증권",
        "aliases": [
          "유진투자증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "yuhwa",
        "name": "유화증권",
        "aliases": [
          "유화증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "im",
        "name": "iM증권",
        "aliases": [
          "iM증권",
          "아이엠증권",
          "하이투자증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "ibk",
        "name": "IBK투자증권",
        "aliases": [
          "IBK투자증권",
          "아이비케이투자증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "kakaopay",
        "name": "카카오페이증권",
        "aliases": [
          "카카오페이증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "kb",
        "name": "KB증권",
        "aliases": [
          "KB증권",
          "케이비증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "kr",
        "name": "KR투자증권",
        "aliases": [
          "KR투자증권",
          "케이알투자증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "cape",
        "name": "케이프투자증권",
        "aliases": [
          "케이프투자증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "koreaasset",
        "name": "코리아에셋투자증권",
        "aliases": [
          "코리아에셋투자증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "kiwoom",
        "name": "키움증권",
        "aliases": [
          "키움증권"
        ],
        "profiles": [
          {
            "columns": {
              "dateCol": [
                "거래일자"
              ],
              "tradeTypeCol": [
                "적요명"
              ],
              "amountCol": [
                "거래금액(외)"
              ],
              "balanceCol": [
                "외화예수금잔고"
              ],
              "feeCol": [
                "수수료(외)"
              ],
              "taxCol": [
                "외국납부세액"
              ],
              "stockNameCol": [
                "종목명"
              ],
              "stockQtyCol": [
                "거래수량"
              ],
              "unitPriceCol": [
                "거래단가/환율",
                "‘거래단가/환율"
              ],
              "stockBalanceCol": [
                "유가금잔"
              ]
            },
            "depositTypes": [
              "매도",
              "배당금(외화)입금",
              "외국납부세환급(외화)입금",
              "외화예탁금이용료(이자)입금"
            ],
            "withdrawalTypes": [
              "매수",
              "외국납부세출금(외화)",
              "외화매도"
            ]
          }
        ],
        "shared": {
          "depositTypes": [
            "매도",
            "배당금(외화)입금",
            "외국납부세환급(외화)입금",
            "외화예탁금이용료(이자)입금"
          ],
          "withdrawalTypes": [
            "매수",
            "외국납부세출금(외화)",
            "외화매도"
          ],
          "removalTypes": []
        },
        "exceptions": [
          {
            "id": "kiwoom-foreign-krw-usd-csv",
            "label": "키움 해외 원화·외화 통합 거래내역(통화 열 포함)",
            "cases": 1,
            "source": "cases/20261110_키움해외_요약.txt",
            "signature": [
              "거래일자",
              "거래종류",
              "통화",
              "거래수량",
              "거래금액",
              "정산금액",
              "세금합",
              "인지세",
              "미수(원/주)",
              "변제합",
              "매체구분",
              "종목코드",
              "적요명",
              "거래단가/환율",
              "예수금잔고",
              "외화예수금잔고",
              "유가금잔",
              "연체합",
              "처리시간",
              "거래소",
              "종목명",
              "거래금액(외)",
              "정산금액(외)",
              "수수료(외)",
              "미수(외)",
              "변제합(외)",
              "외국납부세액"
            ],
            "classificationBasis": "balance",
            "currencySplit": {
              "currencyCol": [
                "통화"
              ],
              "baseCurrency": "KRW",
              "base": {
                "columns": {
                  "amountCol": [
                    "거래금액"
                  ],
                  "balanceCol": [
                    "예수금잔고"
                  ]
                }
              },
              "foreign": {
                "columns": {
                  "settlementCol": [
                    "정산금액(외)"
                  ],
                  "extraCostCol1": [
                    "인지세"
                  ]
                }
              }
            }
          }
        ]
      },
      {
        "id": "toss",
        "name": "토스증권",
        "aliases": [
          "토스증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "hana",
        "name": "하나증권",
        "aliases": [
          "하나증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "sc",
        "name": "한국스탠다드차타드증권",
        "aliases": [
          "한국스탠다드차타드증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "koreainvest",
        "name": "한국투자증권",
        "aliases": [
          "한국투자증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "hanyang",
        "name": "한양증권",
        "aliases": [
          "한양증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "hanhwa",
        "name": "한화투자증권",
        "aliases": [
          "한화투자증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "hyundai",
        "name": "현대차증권",
        "aliases": [
          "현대차증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      },
      {
        "id": "heungkuk",
        "name": "흥국증권",
        "aliases": [
          "흥국증권"
        ],
        "profiles": [],
        "shared": {
          "depositTypes": [],
          "withdrawalTypes": [],
          "removalTypes": []
        }
      }
    ],
    "keywordRules": {
      "directionHints": {
        "deposit": [
          "매도",
          "입금",
          "배당",
          "환급",
          "입고",
          "이입"
        ],
        "withdrawal": [
          "매수",
          "출금",
          "출고",
          "수수료",
          "세금",
          "제출",
          "이체"
        ]
      },
      "removalHints": [
        "수수료",
        "세금",
        "제세금",
        "기타수수료",
        "기타제세금",
        "이체",
        "대체",
        "정산",
        "환전",
        "출고",
        "연체료"
      ],
      "blockDeposit": [
        "수수료",
        "세금",
        "제세금",
        "기타수수료",
        "기타제세금",
        "출금",
        "출고",
        "연체료",
        "대출이자"
      ],
      "mirrorPatterns": [
        {
          "name": "sell_summary_to_detail",
          "detailKeywords": [
            "주식매도출고",
            "매도출고"
          ],
          "summaryKeywords": [
            "주식매도입금",
            "매도입금"
          ],
          "detailCategory": "deposit",
          "summaryCategory": "remove"
        },
        {
          "name": "buy_summary_to_detail",
          "detailKeywords": [
            "주식매수입고",
            "매수입고"
          ],
          "summaryKeywords": [
            "주식매수출금",
            "매수출금"
          ],
          "detailCategory": "withdrawal",
          "summaryCategory": "remove"
        }
      ]
    }
  },
  "schemaVersion": 2,
  "changelog": [
    {
      "version": 2,
      "at": "2026-11-10T11:28:58.011Z",
      "changes": [
        "해외 키움증권 형식1: 분류 이름 오타 수정(외화납부세환급(외화)입금 → 외국납부세환급(외화)입금)",
        "해외 키움증권 형식1: 통화 분리(currencySplit) 추가 — 통화 열로 원화(거래금액·예수금잔고)/외화(거래금액(외)·외화예수금잔고) 계좌 분리, 원화 쪽 분류, 외화 정산금액(외)로 숨은 비용 계산"
      ]
    },
    {
      "version": 3,
      "at": "2026-11-10T11:40:46.196Z",
      "changes": [
        "해외 키움증권 형식1: 외화 쪽 비용 열에 \"인지세\" 지정(currencySplit.foreign.columns.extraCostCol1) — 매도 시 인지세가 \"(인지세)\" 출금 행으로 표시됨"
      ]
    },
    {
      "version": 4,
      "at": "2026-11-10T12:28:57.091Z",
      "changes": [
        "해외 키움증권: 사례 1건에서 나온 규칙(통화 분리·인지세·정산금액 열)을 형식1에서 예외 규칙(exceptions)으로 이동 — 열 이름 27개가 모두 같은 파일에만 적용, 입금/출금은 잔액 기준 분류(원화 쪽 고정 분류 목록 삭제)"
      ]
    },
    {
      "version": 5,
      "at": "2026-11-11T00:58:28.662Z",
      "changes": [
        "은행 이름 목록(bank.names) 21곳 추가 — 편집 화면 기관 선택 목록용(파일명 약칭·은행 형식에서 모음)"
      ]
    },
    {
      "version": 6,
      "at": "2026-11-11T01:04:44.009Z",
      "changes": [
        "국내 미래에셋증권: 새 파일 형식 \"홈페이지 거래내역 간단히 보기(웹페이지)\"(형식2, 열 7개) 추가",
        "국내 미래에셋증권: 거래종류 분류 갱신 (입금 2 · 출금 1 · 제거 2)",
        "국내 미래에셋증권: 증권사 공통 분류 갱신"
      ]
    },
    {
      "version": 7,
      "at": "2026-11-12T08:27:55.846Z",
      "changes": [
        "국내 iM증권: 새 파일 형식 \"거래내역(잔액 없음)\"(형식1, 열 9개) 추가",
        "국내 iM증권: 거래종류 분류 갱신 (입금 1 · 출금 1 · 제거 0)",
        "국내 iM증권: 증권사 공통 분류 갱신",
        "국내 iM증권: 수수료 처리 → 거래금액에 포함",
        "약칭 \"아이엠증권·아이엠\" → iM증권"
      ]
    },
    {
      "version": 8,
      "at": "2026-11-12T08:47:27.195Z",
      "changes": [
        "국내 KB증권: 거래종류 분류 갱신 (입금 1 · 출금 2 · 제거 2)",
        "국내 KB증권: 증권사 공통 분류 갱신",
        "국내 KB증권: 수수료 처리 → 거래금액에 포함"
      ]
    },
    {
      "version": 9,
      "at": "2026-11-12T08:48:59.313Z",
      "changes": [
        "국내 대신증권: 추가 적요 거래상대명 + 비고",
        "국내 대신증권: 거래종류 분류 갱신 (입금 4 · 출금 6 · 제거 5)",
        "국내 대신증권: 증권사 공통 분류 갱신"
      ]
    },
    {
      "version": 10,
      "at": "2026-11-12T08:49:35.660Z",
      "changes": [
        "신한은행: 새 은행 파일 형식 추가",
        "신한은행: 적요 열 내용 + 적요 + 비고, 제외 값 8개"
      ]
    },
    {
      "version": 11,
      "at": "2026-11-12T08:51:44.508Z",
      "changes": [
        "국내 한화투자증권: 새 파일 형식(형식3, 열 10개) 추가",
        "국내 한화투자증권: 추가 적요 비고",
        "국내 한화투자증권: 거래종류 분류 갱신 (입금 0 · 출금 0 · 제거 4)",
        "국내 한화투자증권: 증권사 공통 분류 갱신"
      ]
    },
    {
      "version": 12,
      "at": "2026-11-12T08:51:57.232Z",
      "changes": [
        "국내 한화투자증권: 거래종류 분류 갱신 (입금 1 · 출금 1 · 제거 2)",
        "국내 한화투자증권: 증권사 공통 분류 갱신 (바뀐 분류: [입금] 은행이체입금(…)(제거→입금), [출금] 청약증거금출금(제거→출금))"
      ]
    },
    {
      "version": 13,
      "at": "2026-11-12T08:52:37.241Z",
      "changes": [
        "국내 미래에셋증권: 추가 적요 내계좌표시 + 비고",
        "국내 미래에셋증권: 거래종류 분류 갱신 (입금 3 · 출금 2 · 제거 4)"
      ]
    },
    {
      "version": 14,
      "at": "2026-11-12T08:53:34.230Z",
      "changes": [
        "국내 신한투자증권: 추가 적요 의뢰자명",
        "국내 신한투자증권: 거래종류 분류 갱신 (입금 5 · 출금 3 · 제거 2)"
      ]
    },
    {
      "version": 15,
      "at": "2026-11-12T08:54:12.314Z",
      "changes": [
        "국내 iM증권: 추가 적요 비고"
      ]
    },
    {
      "version": 16,
      "at": "2026-11-12T08:54:29.331Z",
      "changes": [
        "국내 한국투자증권: 추가 적요 비고",
        "국내 한국투자증권: 거래종류 분류 갱신 (입금 4 · 출금 4 · 제거 2)"
      ]
    },
    {
      "version": 17,
      "at": "2026-11-12T08:54:54.839Z",
      "changes": [
        "국내 KB증권: 추가 적요 비고"
      ]
    },
    {
      "version": 18,
      "at": "2026-11-12T08:55:31.079Z",
      "changes": [
        "국내 LS증권: 추가 적요 상대계좌명 + 비고"
      ]
    },
    {
      "version": 19,
      "at": "2026-11-12T08:56:04.437Z",
      "changes": [
        "국내 NH투자증권: 추가 적요 받는통장 표시내용 + 비고",
        "국내 NH투자증권: 거래종류 분류 갱신 (입금 3 · 출금 3 · 제거 2)"
      ]
    },
    {
      "version": 20,
      "at": "2026-11-12T08:57:55.593Z",
      "changes": [
        "국내 SK증권: 새 파일 형식(형식2, 열 10개) 추가",
        "국내 SK증권: 추가 적요 비고",
        "국내 SK증권: 거래종류 분류 갱신 (입금 1 · 출금 2 · 제거 1)",
        "국내 SK증권: 증권사 공통 분류 갱신",
        "국내 SK증권: 수수료 처리 → 거래금액에 포함"
      ]
    },
    {
      "version": 21,
      "at": "2026-11-12T09:12:06.263Z",
      "changes": [
        "거래처 이름이 붙은 거래종류를 앞부분 일치로 변경(사전에 거래처 이름을 남기지 않음, 다른 거래처에도 같은 분류 적용)",
        "국내 한화투자증권 형식3: \"[입금] 은행이체입금(거래처)\" → \"[입금] 은행이체입금(…)\"",
        "국내 한화투자증권 형식3: \"[출고] 타사대체출고(거래처)\" → \"[출고] 타사대체출고(…)\"",
        "국내 한화투자증권 공통 분류: \"[입금] 은행이체입금(거래처)\" → \"[입금] 은행이체입금(…)\"",
        "국내 한화투자증권 공통 분류: \"[출고] 타사대체출고(거래처)\" → \"[출고] 타사대체출고(…)\""
      ]
    },
    {
      "version": 22,
      "at": "2026-11-12T09:15:11.239Z",
      "changes": [
        "변경 기록 안의 거래처 이름 1곳을 (…)로 가림"
      ]
    }
  ]
}; })(typeof window !== 'undefined' ? window : globalThis);
