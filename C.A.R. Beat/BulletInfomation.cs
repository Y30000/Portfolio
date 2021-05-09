using UnityEngine;

public class BulletInfomation : MonoBehaviour
{
    public float InfoRotateAnglePerSec;
    public float InfoMoveSpeed;
    public float InfoBulletLifeTime;
    public float InfoTraceRange;

    public string InfoTargetTag;
    public Vector3 InfoFirstLookDirection;

    public float Damage;
    [HideInInspector]
    public Rigidbody rb;

    private void Awake()
    {
        GameObject player = GameObject.Find("Player");

        if(player == null)
        {
            Destroy(gameObject);
            return;
        }
        rb = GetComponent<Rigidbody>();
        if (rb == null)
        {
            rb = gameObject.AddComponent<Rigidbody>();
            rb.useGravity = false;
            rb.drag = 0;
            rb.angularDrag = 0;
        }
    }

    public void Explosion()
    {
        print("폭발 이팩트 생성 여기");
        Destroy(gameObject);
    }
}