# BulletInfomation.cs

# 투사체와 투사체의 target에 대한 기본 정보를 담고 있음


# ********************


# GuidedMissile.cs

# BulletInfomation 으로부터 투사체의 특성중 운동에 대한특성만을 가져오고, 투사체에 붙어있는 Sphere Collider의 Is Trigger 옵션을 켜두고 탐지범위와 target에 대한 tag를 설정 이후 Sphere Collider 내에 들어온 target tag를 가진 개채를 추적. 또한, target인 개체가 최소 회전 반경 안쪽으로 들어오면 맞출 수 없도록 설계됨