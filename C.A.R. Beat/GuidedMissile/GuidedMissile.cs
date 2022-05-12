using System.Collections.Generic;

public class GuidedMissile : MonoBehaviour
{
    private List<Transform> targets;
    private Rigidbody rb;

    private float rotateAnglePerSec;
    private float moveSpeed;

    private string targetTag;
    private Vector3 firstLookDirection;
    private Transform thisTransform;

    // Start is called before the first frame update
    private void Start()
    {
        thisTransform = transform;

        targets = new List<Transform>();

        if (firstLookDirection == Vector3.zero)
        {
            firstLookDirection = thisTransform.right;                       //시작시 초기화 방향을 지정하지 않았다면 오른쪽을 보도록
        }

        BulletInfomation info = GetComponentInParent<BulletInfomation>(); //초당 화전 각, 속력, 타겟의 테그, 초기화시 자신이 바라보고 있을 방향이 들어있는 컴포넌트

        if (info == null)
        {
            Debug.LogError("Error Don't have Infomation");
        }

        rb = info.rb;

        rotateAnglePerSec = info.InfoRotateAnglePerSec;
        moveSpeed = info.InfoMoveSpeed;
        targetTag = info.InfoTargetTag;
        firstLookDirection = info.InfoFirstLookDirection;

        if (targetTag == EnumTags.Player.ToString())
        {
            targets.Add(GameObject.Find("Player").transform);
        }

        Enemy_Guider_Init(firstLookDirection, moveSpeed);
    }

    private void Enemy_Guider_Init(Vector3 initDirection, float initSpeed)      //외부에서 다시 호출할 때 도 사용
    {
        moveSpeed = initSpeed;
        rb.velocity = thisTransform.forward * moveSpeed;
    }

    private void Update()
    {
        Vector3 dir;
        Transform target = null;

        if (targets.Count > 0)  //타겟으로 등록된 물체가 있으면 실행
        {
            for (int i = targets.Count; i > 0;)
            {
                Transform t = targets[--i];
                if (target == null)
                {
                    target = t;
                    continue;
                }

                if (Vector3.Distance(thisTransform.position, target.position) > Vector3.Distance(thisTransform.position, t.position))   //가장 가까운 타겟 저장
                {
                    target = t;
                }
            }
        }

        if (null == target)
        {
            dir = rb.velocity;  //타겟이 없다면 현재 진행방향
        }
        else
        {
            dir = target.position - thisTransform.position; //타겟이 있다면 타겟 방향
        }

        dir = dir.normalized;   //방향만 남기고 힘은 제거
        float degree = -Mathf.Atan2(dir.z, dir.x) * Mathf.Rad2Deg + 90; //자신과 타겟과의 벡터를 각도로 변환 +90의 이유는 Vector3(0,0,1) 이 0도인 값이 나오기 때문

        float RAPS = rotateAnglePerSec * Time.deltaTime;

        if (target != null) //타겟이 존재할 때
        {
            float angle = Vector3.Angle(thisTransform.forward, target.position - thisTransform.position);

            thisTransform.rotation = Quaternion.Slerp(thisTransform.rotation, Quaternion.Euler(0, degree, 0), ((RAPS > angle) ? 1 : RAPS / angle));     //타겟을 향해 자신 회전, lerp를 사용해서 균등한 속도로 회전하도록 함
            rb.velocity = thisTransform.forward * moveSpeed;                                                        //자신의 속도를 바라보고있는 방향과, 일정한 힘으로 변경
            Debug.DrawRay(thisTransform.position, dir, Color.magenta);                                              //타겟의 방향을 확인하기 위한 디버그용 레이
        }
        else                //타겟이 존재하지 않을 때
        {
            thisTransform.rotation = Quaternion.Slerp(thisTransform.rotation, Quaternion.Euler(0, degree, 0), 1);   //자신의 이동방향을 정면으로 보도록 회전
            rb.velocity = thisTransform.forward * rb.velocity.magnitude;
        }

        if (rotateAnglePerSec == 0)     //직진만 하는 적의 경우 한 프레임만 회전한 후 컴포넌트 비활성화
        {
            enabled = false;
        }
    }

    private void OnDisable()
    {

        if (rotateAnglePerSec == 0)     //회전이 없는 물체라도 처음 방향은 필요하기에 오브젝트 풀링을 위해 다시 활성화
        {
            enabled = true;
        }
    }

    private void GuidOver()                             //스테이지나 게임이 끝날때 호출
    {
        targets.Clear();
        Destroy(gameObject.GetComponent<SphereCollider>());
        rb.useGravity = true;
    }

    //////////////////////////////////////////////////////////////  targets 관리 시작
    private void OnTriggerEnter(Collider other)
    {
        if (other.gameObject.CompareTag(targetTag))     //충돌체의 테그가 지정된 테그가 맞다면 타겟 등록
        {
            targets.Add(other.gameObject.transform);
        }
    }

    private void OnTriggerExit(Collider other)
    {
        if (targets.Contains(other.gameObject.transform)) //충돌체의 테그가 지정된 테그가 맞다면 타겟 해제
        {
            targets.Remove(other.gameObject.transform);

            if (targets.Count == 0)                         //타겟이 존재하지 않으면 존재할 이유가 사라짐
            {
                Destroy(thisTransform.parent.gameObject);
            }
        }

        
    }
    //////////////////////////////////////////////////////////////  targets 관리 끝
}