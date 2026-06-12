# -*- coding: utf-8 -*-
"""EMU(English Metric Unit) 변환 상수의 단일 진실원.

PowerPoint geometry는 EMU 단위(1cm = 360000 EMU, 1pt = 12700 EMU)를 쓴다.
이 두 상수가 여러 모듈에 흩어져 있으면 값이 갈릴(드리프트) 위험이 있어 한 곳에 모은다.
의존성이 없는 leaf 모듈이라 어느 모듈에서 import해도 순환이 생기지 않는다.
값은 float로 통일(나눗셈·임계 비교에 그대로 쓰고, geometry 대입은 호출부에서 int()/Emu()로 감싼다).
"""
EMU_PER_CM = 360000.0
EMU_PER_PT = 12700.0
