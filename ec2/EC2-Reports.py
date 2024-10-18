import boto3
import pandas as pd

# Get the profile name and regions as input
profilename = input('Enter the name of profile: ')
regions_input = input('Enter the region ailas (comma-separated if multiple): ')
save_path = input('Enter the path(optional) and the Name of the file to save without extension: ')

# Split the input regions into a list
regions = regions_input.split(',')

with pd.ExcelWriter(save_path + '.xlsx', engine='xlsxwriter') as writer:
    for region_name in regions:
        # Set up the EC2 client for the current region
        session = boto3.Session(profile_name=profilename, region_name=region_name.strip())  # Strip to remove leading/trailing spaces
        ec2 = session.client('ec2')

        # Initialize a DataFrame to store data for the current region
        data = []

        for reservation in ec2.describe_instances()['Reservations']:
            for instance in reservation['Instances']:
                instance_id = instance['InstanceId']
                public_ip = instance.get('PublicIpAddress', '')
                private_ip = instance.get('PrivateIpAddress', '')
                instance_type = instance.get('InstanceType', '')
                instance_state = instance['State']['Name']
                key_name = instance.get('KeyName', '')
                subnet = instance.get('SubnetId', '')
                vpc = instance.get('VpcId', '')
                instance_profile = instance.get('IamInstanceProfile', {}).get('Arn', '')

                tags = instance.get('Tags', [])

                # Retrieve values for specified tags
                name = next((tag['Value'] for tag in tags if tag['Key'] == 'Name'), '')
                owner = next((tag['Value'] for tag in tags if tag['Key'] == 'Owner'), '')
                operating_system = next((tag['Value'] for tag in tags if tag['Key'] == 'Operating System'), '')
                purpose = next((tag['Value'] for tag in tags if tag['Key'] == 'Purpose'), '')
                product = next((tag['Value'] for tag in tags if tag['Key'] == 'Product'), '')
                component = next((tag['Value'] for tag in tags if tag['Key'] == 'Component'), '')
                nature = next((tag['Value'] for tag in tags if tag['Key'] == 'Nature'), '')
                patch_group = next((tag['Value'] for tag in tags if tag['Key'] == 'Patch Group'), '')
                priority = next((tag['Value'] for tag in tags if tag['Key'] == 'Maintenance Windows'), '')
                asg = next((tag['Value'] for tag in tags if tag['Key'] == 'aws:autoscaling:groupName'), '')
                stack = next((tag['Value'] for tag in tags if tag['Key'] == 'aws:cloudformation:stack-id'), '')
                fleet = next((tag['Value'] for tag in tags if tag['Key'] == 'aws:ec2:fleet-id'), '')
                ebs = next((tag['Value'] for tag in tags if tag['Key'] == 'elasticbeanstalk:environment-id'), '')
                spot = next((tag['Value'] for tag in tags if tag['Key'] == 'Schedule'), '')

                # Get attached volumes
                volumes = [block_device.get('Ebs', {}).get('VolumeId', '') for block_device in instance.get('BlockDeviceMappings', [])]

                # Get security groups
                security_groups = instance.get('SecurityGroups', [])
                security_group_ids = [group['GroupId'] for group in security_groups]
                security_group_names = [group['GroupName'] for group in security_groups]

                data.append([instance_id, public_ip, private_ip, instance_type, instance_state, key_name, name, owner, operating_system,
                             purpose, product, component, nature, patch_group, priority,
                             instance_profile, asg, stack, fleet, ebs, spot, ', '.join(security_group_ids), ', '.join(security_group_names), ', '.join(volumes), subnet, vpc])

        # Create a DataFrame for the current region's data
        df = pd.DataFrame(data, columns=['Instance ID', 'Public IP', 'Private IP', 'InstanceType', 'InstanceState', 'Keyname', 'Name', 'Owner', 'Operating System', 'Purpose', 'Product', 'Component', 'Nature', 'Patch Group', 'Priority', 'Instance Profile', 'ASG', 'STACK', 'Fleet', 'EBS', 'SPOT', 'Security Groups ID', 'Security Groups Name', 'Volumes', 'SubnetID', 'VPCID'])

        # Write the DataFrame to the Excel file with the region name as the sheet name
        df.to_excel(writer, sheet_name=region_name, index=False)

print('Excel report has been generated and saved at ' + save_path + '.xlsx')
