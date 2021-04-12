import * as React from 'react';
import * as strings from 'MyTodoWebPartStrings';
import styles from './MyTodo.module.scss';
import { IMyTodoProps } from './IMyTodoProps';
import { GraphRequest } from '@microsoft/sp-http';
import { List, Link, Label } from 'office-ui-fabric-react';
import { IMyTodoState } from './IMyTodoState';
import { ITaskList, ITasksLists } from './ITaskList';
import { ITasks, ITask } from './ITasks';
import { FontIcon, IIconProps } from 'office-ui-fabric-react/lib/Icon';
import { IconButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';
import { mergeStyles, mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/components/Spinner';
import { Separator } from 'office-ui-fabric-react/lib/Separator';

export default class MyTodo extends React.Component<IMyTodoProps, IMyTodoState> {

  private _interval: number;

  private iconClass = mergeStyles({
    fontSize: 24,
    height: 24,
    width: 24,
    margin: '2px 5px',
    color: 'blue'
  });

  private _activeTaskIcon: IIconProps = { iconName: 'StatusCircleRing' };
  private _completedTaskIcon: IIconProps = { iconName: 'SkypeCircleCheck' };

  constructor(props: IMyTodoProps) {
    super(props);
    this.state = {
      tasks: [],
      activeTaskList: null,
      loading: false,
      error: undefined,
      renderedDateTime: new Date(),
    };
  }

  private _getTaskList = () => {
    if (!this.props.graphClient) {
      return;
    }
    const request: GraphRequest = this.props.graphClient
      .api("me/todo/lists")
      .version("v1.0");
    return request.get((err: any, res: ITasksLists): void => {

      // Check if a response was retrieved
      if (err) {
        console.error(err);
      }
      else if (res && res.value && res.value.length > 0) {
        const myLists: ITaskList[] = res.value;
        // filter through task list since issues with graph doing filter
        const activeTaskList: ITaskList = (myLists.filter(list => list.wellknownListName === "defaultList"))[0];
        this._getTasks(activeTaskList);
      }
      return null;
    });
  }

  private _getTasks = (defaultTaskList: ITaskList) => {
    if (!this.props.graphClient) {
      return;
    }

    this.setState({
      loading: true
    });

    const apiUrl = "me/todo/lists/" + defaultTaskList.id + "/tasks";

    const request: GraphRequest = this.props.graphClient
      .api(apiUrl)
      .filter(`status ne 'completed'`)
      .top(this.props.nrOfTasks || 5)
      .version("v1.0");
    return request.get((err: any, res: ITasks): void => {
      // Check if a response was retrieved
      if (err) {
        console.error("Error: " + err);
        this.setState({
          error: err.message ? err.message : strings.Error,
          loading: false
        });
      }
      else if (res && res.value && res.value.length > 0) {
        const myTasks: ITask[] = res.value;
        this.setState({
          tasks: myTasks,
          loading: false,
          activeTaskList: defaultTaskList
        });
      }
    });
  }

  private _setInterval = (): void => {
    let { refreshInterval } = this.props;
    // set up safe default if the specified interval is not a number
    // or beyond the valid range
    if (isNaN(refreshInterval) || refreshInterval < 0 || refreshInterval > 60) {
      refreshInterval = 5;
    }
    // refresh the component every x minutes
    this._interval = setInterval(this._reRender, refreshInterval * 1000 * 60);
    this._reRender();
  }

  private _reRender = (): void => {
    this._getTaskList();
    this.setState({ renderedDateTime: new Date() });
  }

  public componentDidMount(): void {
    this._setInterval();
  }

  public componentWillUnmount(): void {
    // remove the interval so that the data won't be reloaded
    clearInterval(this._interval);
  }

  public componentDidUpdate(prevProps: IMyTodoProps, prevState: IMyTodoState): void {

    if (prevProps.refreshInterval !== this.props.refreshInterval) {
      clearInterval(this._interval);
      this._setInterval();
      return;
    }

    // verify if the component should update. Helps avoid unnecessary re-renders
    // when the parent has changed but this component hasn't
    if ((prevProps.nrOfTasks !== this.props.nrOfTasks) ||
      (prevState.renderedDateTime !== this.state.renderedDateTime)) {
      this._getTaskList();
    }
  }

  private _changeTaskStatus = (taskId: string, status: string) => {
    if (!this.props.graphClient) {
      return;
    }

    const graphRequest: GraphRequest = this.props.graphClient
      .api(`me/todo/lists/${this.state.activeTaskList.id}/tasks/${taskId}`)
      .version("v1.0");

    graphRequest.patch({ status: `${status}` });

    const tasks: ITask[] = [];
    this.state.tasks.forEach((task: ITask) => {
      if (task.id !== taskId) {
        tasks.push(task);
      }
    });

    this.setState({
      tasks: tasks
    });
  }

  private _changeTaskImportance = (taskId: string, importance: string) => {
    if (!this.props.graphClient) {
      return;
    }

    importance = importance == 'normal' ? 'high' : 'normal';
    console.log("Importance" + importance);

    const graphRequest: GraphRequest = this.props.graphClient
      .api(`me/todo/lists/${this.state.activeTaskList.id}/tasks/${taskId}`)
      .version("v1.0");

    graphRequest.patch({ importance: `${importance}` });

    const tasks: ITask[] = [];
    this.state.tasks.forEach((task: ITask) => {
      if (task.id == taskId) {
        task.importance = importance;
      }
      tasks.push(task);
    });

    this.setState({
      tasks: tasks
    });
  }

  private _onRenderCell = (item: ITask, index: number | undefined): JSX.Element => {
    const iconClass = mergeStyles({
      fontSize: 24,
      height: 24,
      width: 24,
      margin: '2px 5px',
      color: 'blue',
      float:'left'
   });

    const iconImportance = mergeStyles({
      fontSize: 24,
      height: 24,
      width: 24,
      margin: '2px 5px',
      color: 'blue',
    });

    const activeTaskIcon: IIconProps = {
      iconName: 'StatusCircleRing',
      className: iconClass
    };

    const highImportanceIcon: IIconProps = {
      iconName: 'FavoriteStarFill',
      className: iconImportance
    };

    const normalImportanceIcon: IIconProps = {
      iconName: 'FavoriteStar',
      className: iconImportance
    };

    const completedTaskIcon: IIconProps = {
      iconName: 'SkypeCircleCheck',
      className: iconClass
    };

    const dateDisplay = {
      text: "",
      className: styles.date
    };

    if (item.dueDateTime) {
      const due = new Date(item.dueDateTime.dateTime);
      due.setHours(0,0,0,0);
      
      const today = new Date();
      today.setHours(0,0,0,0);      

      if (due < today) {
        dateDisplay.text = `Overdue, ${due.toLocaleDateString()}`;
        dateDisplay.className += ` ${styles.overdue}`;
      }
      else if (due.getTime() === today.getTime()) {
        dateDisplay.text = "Due Today";
        dateDisplay.className += ` ${styles.dueToday}`;
      }
      else if (due.getTime() === today.getTime() + (1000 * 60 * 60 * 24)) {
        dateDisplay.text = "Due Tomorrow";
      }
      else {
        dateDisplay.text = `Due ${due.toLocaleDateString()}`;
      }

    }

    return (
      <div className={styles.todoItem}>
        {
          (index > 0) &&
          <Separator />
        }
        <IconButton
          className={styles.todoStatus}
          iconProps={activeTaskIcon}
          title={`Complete ${item.title}`}
          ariaLabel={`Complete ${item.title}`}
          disabled={false}
          checked={false}
          data={item.id}
          onClick={() => this._changeTaskStatus(item.id, "completed")} />
        <div className={styles.todoDetails}>
          <Link href={`https://to-do.office.com/tasks/id/${item.id}/details`} target="_blank" className={styles.todoLink}>{item.title}</Link>
          {
            item.dueDateTime &&
            <div className={dateDisplay.className} >{dateDisplay.text}</div>
          }
        </div>
        <IconButton
          className={styles.todoImportance}
          iconProps={item.importance == 'normal' ? normalImportanceIcon : highImportanceIcon}
          title={item.importance == 'normal' ? `Mark task important: ${item.title}` : 'Remove importance'}
          ariaLabel={item.importance == 'normal' ? `Mark task important: ${item.title}` : 'Remove importance'}
          disabled={false}
          checked={false}
          data={item.id}
          onClick={() => this._changeTaskImportance(item.id, item.importance)} />          
      </div>
    );
  }


  public render(): React.ReactElement<IMyTodoProps> {
    return (
      <div className={styles.todo} >
        <WebPartTitle displayMode={this.props.displayMode}
          title={this.props.title}
          updateProperty={this.props.updateProperty} className={styles.title} />
        {
          this.state.loading &&
          <Spinner label={strings.Loading} size={SpinnerSize.large} />
        }
        {
          this.state.tasks &&
            this.state.tasks.length > 0 ? (
            <Stack>
              <Stack className={styles.list}>
                <List items={this.state.tasks}
                  onRenderCell={this._onRenderCell} />
              </Stack>
              <Link href='https://to-do.office.com/tasks/' target='_blank' className={styles.viewAll}>{strings.ViewAllTodo}</Link>
            </Stack>
          ) : (
            !this.state.loading && (
              this.state.error ?
                <span >{this.state.error}</span> :
                <span ></span>
            )
          )
        }
      </div >
    );
  }
}
