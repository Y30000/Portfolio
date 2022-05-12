using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class TurretAiming : MonoBehaviour
{
    public Transform Pool;
    public List<Transform> TargetPools;
    public string TagName = "Player";
    public float rotatePerSec = 10f;

    public float targetMaximumRange = 15f;

    public Transform fixedTarget;
    public Transform targetTransform;
    private Transform thisTransform;
    private List<Transform> bosss;

    public LayerMask LayerMask;
    public float delayTime = 0f;

    public bool autoRetargeting = false;
    private readonly float bossBonuce = 5f;
    private Transform pools;

    // Start is called before the first frame update
    private void Start()
    {
        thisTransform = transform;
        pools = GameObject.Find("Pools").transform;
        bosss = new List<Transform>
        {
            GameObject.Find("BossSpawner").transform.Find("BBoss"),                                   //bboss
            GameObject.Find("BossSpawner").transform.Find("BossPlanSBoss").transform.Find("SBoss")    //sboss
        };

    }

    private bool isDelay = false;

    private IEnumerator DelayTimmer()
    {
        float time = delayTime;
        float delay = 0.5f;
        isDelay = true;
        while (time > 0)
        {
            time -= delay;
            yield return new WaitForSeconds(delay);
        }
        isDelay = false;
    }

    // Update is called once per frame
    private void FixedUpdate()
    {
        if (isDelay)
        {
            return;
        }

        if (targetTransform)
        {
            if (!targetTransform.gameObject.activeSelf && delayTime != 0)
            {
                StartCoroutine(DelayTimmer());
                targetTransform = null;
                return;
            }
        }

        if(null == targetTransform || !targetTransform.gameObject.activeSelf || autoRetargeting)
        {
            if (TagName == "Player")
            {
                try
                {
                    targetTransform = GameObject.Find(TagName).GetComponent<Transform>();
                }
                catch
                {
                    targetTransform = null;
                }
            }
            else if (TagName == "Enemy")
            {
                float distance = float.MaxValue;

                if (TargetPools.Count == 0)
                {
                    for(int i = pools.childCount; i > 0;)
                    {
                        Transform child = pools.GetChild(--i);
                        if (child.name.Contains("Enemy_"))
                        {
                            TargetPools.Add(child);
                        }
                    }
                }

                targetTransform = null;
                //foreach(var target in Physics.OverlapSphere(transform.position, targetMaximumRange, LayerMask, QueryTriggerInteraction.Collide))
                for(int i = TargetPools.Count; i > 0; )
                {
                    Transform child = TargetPools[--i];
                    for (int j = child.childCount; j > 0;)
                    {
                        Transform target = child.GetChild(--j);
                        if (!target.gameObject.activeSelf)
                        {
                            continue;
                        }

                        float newDistance = Vector3.Distance(thisTransform.position, target.transform.position) - ((target.name.Contains("Boss")) ? bossBonuce : 0);

                        if (distance > newDistance && target.tag == TagName)
                        {
                            targetTransform = target.transform;
                            distance = newDistance;
                        }
                    }
                }

                for(int i = bosss.Count; i > 0;)
                {
                    Transform target = bosss[--i];
                    if (!target.gameObject.activeSelf)
                    {
                        continue;
                    }

                    float newDistance = Vector3.Distance(thisTransform.position, target.transform.position) - ((target.name.Contains("Boss")) ? bossBonuce : 0);

                    if (distance > newDistance && target.tag == TagName)
                    {
                        targetTransform = target.transform;
                        distance = newDistance;
                    }
                    
                }

                if (distance > targetMaximumRange)
                {
                    targetTransform = null;
                    return;
                }
            }
            else if (TagName == "Battery")
            {
                float distance = float.MaxValue;

                if (!Pool)
                {
                    return;
                }

                targetTransform = null;
                for (int i = 0; i < Pool.childCount; ++i)
                {
                    if (!Pool.GetChild(i).gameObject.activeSelf || Pool.GetChild(i).GetComponent<Pulled>().IsPulled)
                    {
                        continue;
                    }

                    float newDistance = Vector3.Distance(thisTransform.position, Pool.GetChild(i).position);
                    if (distance > newDistance)
                    {
                        targetTransform = Pool.GetChild(i);
                        distance = newDistance;
                    }

                }

                if (distance > targetMaximumRange)
                {
                    targetTransform = null;
                    return;
                }

                targetTransform.GetComponent<Pulled>().IsPulled = true;
            }
            else if(TagName == "Boss")
            {
                targetTransform = fixedTarget;
            }

            if (null == targetTransform)
            {
                return;
            }
        }

        Vector3 pos = targetTransform.position;
        Vector3 player_pos = thisTransform.position;
        Vector2 target_pos = new Vector2(pos.x - player_pos.x, pos.z - player_pos.z);

        float target_rotate = Mathf.Atan2(target_pos.x, target_pos.y) * Mathf.Rad2Deg;
        float angle = GetAngleDIstance(thisTransform.eulerAngles.y, target_rotate);
        float RPS = rotatePerSec * Time.fixedDeltaTime;
        thisTransform.eulerAngles = new Vector3(0, MyLerp(thisTransform.eulerAngles.y, target_rotate, ((RPS > angle) ? 1 : RPS / angle)) , 0);
    }

    private Vector3 RemainVector(Vector3 value, bool x, bool y, bool z)
    {
        if (!x)
        {
            value.x = 0;
        }

        if (!y)
        {
            value.y = 0;
        }

        if (!z)
        {
            value.z = 0;
        }

        return value.normalized;
    }

    private float MyLerp(float from, float to, float offset)
    {
        if(from < to)
        {
            if(to - from > 180)
            {
                from += 360;
            }
        }
        else
        {
            if(from - to > 180)
            {
                to += 360;
            }
        }
        
        return from + (to - from) * offset;
    }

    private float MyModuler(float value, float max)
    {
        if(value < max)
        {
            if (value < 0) { MyModuler(value, max); }
            else { return value; }
        }
        return MyModuler(value - max, max);
    }

    private float GetAngleDIstance(float from, float to)
    {
        float angle;

        if (to < 0)
        {
            angle = Mathf.Abs(Mathf.Abs(from) - Mathf.Abs(to + 360));
        }
        else
        {
            angle = Mathf.Abs(Mathf.Abs(from) - Mathf.Abs(to));
        }

        if (angle > 180)
        {
            angle -= 360;
            if (angle < 0)
            {
                angle *= -1;
            }
        }

        return angle;
    }
}